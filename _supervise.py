"""
受控看护进程（后台常驻）：
  - 启动/复用独立 Edge（CDP 9223，测试 profile）
  - 清空旧登录态并打开 FIF 登录页，等待用户手动登录（脚本内不含任何账号密码）
  - 登录成功后自动注入刷课脚本
  - 每 3 秒把状态写入 _supervise_status.json
  - 轮询 _supervise_cmd.json 执行远程指令（eval/click/set_input/inject/goto/stop）

指令文件格式（写入后会被读取并删除）：
  {"id":"1","action":"eval","js":"location.href"}
  {"id":"2","action":"click","text":"一键刷全部课程并达标"}
  {"id":"3","action":"set_input","placeholder":"0-100","value":"99"}
  {"id":"4","action":"inject"}
  {"id":"5","action":"goto","url":"https://..."}
  {"id":"6","action":"reload"}
  {"id":"7","action":"stop"}
结果写入 _supervise_result.json
"""
import asyncio
import json
import os
import subprocess
import sys
import time
import traceback
import urllib.request

from playwright.async_api import async_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE = os.path.join(HERE, "_edge_test_profile2")
SCRIPT_FILE = os.path.join(HERE, "course-learning-toolkit.private.user.js")
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
CDP = "http://127.0.0.1:9223"
LOGIN = ("https://www.fifedu.com/iplat/fifLogin/index.html"
         "?service=https%3A%2F%2Ficourse.fifedu.com%2Fistp-manage%2Fwork%2Fhome")
HOME = "https://icourse.fifedu.com/istp-manage/work/home"

STATUS = os.path.join(HERE, "_supervise_status.json")
CMD = os.path.join(HERE, "_supervise_cmd.json")
RESULT = os.path.join(HERE, "_supervise_result.json")
CONSOLE_LOG = os.path.join(HERE, "_supervise_console.log")


def cdp_v():
    try:
        with urllib.request.urlopen(CDP + "/json/version", timeout=2) as r:
            return json.load(r).get("Browser")
    except Exception:
        return None


def ensure_edge():
    if cdp_v():
        return
    subprocess.Popen([EDGE, "--remote-debugging-port=9223", f"--user-data-dir={PROFILE}",
                      "--start-maximized", "--no-first-run", LOGIN], close_fds=True)
    for _ in range(30):
        time.sleep(1)
        if cdp_v():
            return
    sys.exit("CDP 超时")


def write_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


async def main():
    ensure_edge()
    with open(SCRIPT_FILE, "r", encoding="utf-8") as f:
        script = f.read()

    console_lines = []

    async with async_playwright() as p:
        b = await p.chromium.connect_over_cdp(CDP)
        ctx = b.contexts[0]

        def pick_page():
            """挑选活动页面：优先 icourse 学习空间，其次任意 http 页面（登录页等）"""
            pages = [pg for pg in ctx.pages if not pg.is_closed() and pg.url.startswith("http")]
            pref = [pg for pg in pages if "icourse.fifedu.com" in pg.url]
            if pref:
                return pref[-1]
            return pages[-1] if pages else None

        page = pick_page() or await ctx.new_page()

        def push(line):
            console_lines.append(line)
            del console_lines[:-400]
            try:
                with open(CONSOLE_LOG, "a", encoding="utf-8") as f:
                    f.write(line + "\n")
            except Exception:
                pass

        def on_console(msg):
            try:
                push(f"{time.strftime('%H:%M:%S')} [{msg.type}] {msg.text}")
            except Exception:
                pass

        page.on("console", on_console)
        page.on("pageerror", lambda e: push(f"{time.strftime('%H:%M:%S')} [pageerror] {e}"))

        # 登录过程中平台可能新建标签页，新页面同样挂上控制台监听
        def on_new_page(np):
            try:
                np.on("console", on_console)
                np.on("pageerror", lambda e: push(f"{time.strftime('%H:%M:%S')} [pageerror] {e}"))
            except Exception:
                pass

        ctx.on("page", on_new_page)

        # 启动策略：--keep-login 时保留已有登录态（不清理 cookie），直接注入脚本续跑；
        # 否则清登录态等待手动登录。
        keep_login = "--keep-login" in sys.argv
        logged_in = False
        injected = False
        if keep_login:
            print("[SUP] 保留登录态模式，直接打开学习空间 ...")
            await page.goto(HOME, wait_until="domcontentloaded", timeout=45000)
            await asyncio.sleep(3)
            url0 = page.url
            if "fifLogin" in url0 or "login" in url0.lower():
                print("[SUP] 登录态已失效，等待手动登录 ...")
                logged_in = False
            else:
                logged_in = True
                await ctx.add_init_script(script)
                await page.reload(wait_until="domcontentloaded", timeout=45000)
                injected = True
                print("[SUP] 已注入脚本并续跑:", page.url)
                await asyncio.sleep(3)
        else:
            await ctx.clear_cookies()
            await page.goto(LOGIN, wait_until="domcontentloaded", timeout=45000)
            try:
                await page.evaluate("() => { localStorage.clear(); sessionStorage.clear(); }")
                await page.reload(wait_until="domcontentloaded", timeout=45000)
            except Exception:
                pass
            print("[SUP] 已清理旧登录态，等待手动登录 ...")

        loops = 0
        while True:
            loops += 1
            await asyncio.sleep(3)
            # 页面句柄可能因登录跳转/新标签页而失效，每轮重新挑选一次活动页面
            np = pick_page()
            if np is not None and np is not page:
                try:
                    if page.is_closed():
                        print("[SUP] 原页面已关闭，切换页面:", np.url[:100])
                        page = np
                except Exception:
                    page = np
            try:
                url = page.url
            except Exception:
                page = pick_page() or await ctx.new_page()
                url = page.url

            # 登录检测
            if not logged_in:
                if "fifLogin" not in url and "login" not in url.lower():
                    logged_in = True
                    print("[SUP] 检测到已登录:", url)
                    if logged_in:
                        # 注入脚本并 reload
                        await ctx.add_init_script(script)
                        await page.reload(wait_until="domcontentloaded", timeout=45000)
                        injected = True
                        await asyncio.sleep(4)
                        print("[SUP] 脚本已注入")

            # 状态快照（页面句柄失效时换新页面重试一次）
            snap = {"time": time.strftime("%Y-%m-%d %H:%M:%S"), "url": "", "title": "",
                    "panel": "", "logged_in": logged_in, "injected": injected}
            SNAP_JS = """() => ({
                    url: location.href,
                    title: document.title,
                    panel: document.getElementById('fif-helper-panel')?.innerText || '',
                })"""
            try:
                snap.update(await page.evaluate(SNAP_JS))
            except Exception as e:
                try:
                    cand = pick_page()
                    if cand is not None and cand is not page:
                        page = cand
                        snap.update(await page.evaluate(SNAP_JS))
                    else:
                        snap["error"] = str(e)[:200]
                except Exception as e2:
                    snap["error"] = str(e2)[:200]
            snap["console_tail"] = console_lines[-15:]
            write_json(STATUS, snap)

            # 处理指令
            if os.path.exists(CMD):
                try:
                    with open(CMD, "r", encoding="utf-8-sig") as f:
                        cmd = json.load(f)
                    os.remove(CMD)
                except Exception:
                    cmd = None
                    # 解析失败也要删掉，避免同一个坏指令被无限重读
                    try:
                        os.remove(CMD)
                    except Exception:
                        pass
                if cmd:
                    res = {"id": cmd.get("id"), "action": cmd.get("action"), "ok": True}
                    try:
                        a = cmd.get("action")
                        if a == "eval":
                            res["value"] = await page.evaluate(cmd.get("js", "() => null"))
                        elif a == "click":
                            t = cmd.get("text", "")
                            res["value"] = await page.evaluate("""(t) => {
                                const els = Array.from(document.querySelectorAll('button, [role=button], .el-button, a, label, span, div'));
                                const hit = els.filter(x => (x.innerText || '').includes(t) && (x.innerText || '').trim().length < 40);
                                if (!hit.length) return 'not-found';
                                hit[0].click();
                                return hit[0].innerText.trim();
                            }""", t)
                        elif a == "set_input":
                            ph = cmd.get("placeholder", "")
                            val = cmd.get("value", "")
                            res["value"] = await page.evaluate("""([ph, val]) => {
                                const ins = Array.from(document.querySelectorAll('input'));
                                const el = ins.find(i => (i.placeholder || '').includes(ph));
                                if (!el) return 'input-not-found';
                                const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                                setter.call(el, val);
                                el.dispatchEvent(new Event('input', {bubbles: true}));
                                el.dispatchEvent(new Event('change', {bubbles: true}));
                                return el.placeholder + '=' + el.value;
                            }""", [ph, val])
                        elif a == "inject":
                            # 注意：CDP 会话内重复 add_init_script 会让旧版本先注册、新版本被
                            # 「单实例保护」挡住，因此这里只做 reload；脚本更新请重启看护进程
                            await page.reload(wait_until="domcontentloaded", timeout=45000)
                            res["value"] = "reloaded(脚本更新需重启看护进程)"
                        elif a == "reload":
                            await page.reload(wait_until="domcontentloaded", timeout=45000)
                            res["value"] = "reloaded"
                        elif a == "goto":
                            await page.goto(cmd.get("url", HOME), wait_until="domcontentloaded", timeout=45000)
                            res["value"] = page.url
                        elif a == "stop":
                            write_json(RESULT, {**res, "value": "stopped"})
                            print("[SUP] 收到 stop，退出")
                            break
                        else:
                            res["ok"] = False
                            res["value"] = "unknown-action"
                    except Exception as e:
                        res["ok"] = False
                        res["value"] = str(e)[:400]
                    write_json(RESULT, res)
                    print("[SUP] 执行指令:", res)

            if loops % 20 == 0:
                print(f"[SUP] 运行中 {loops*3}s | {url[:90]}")

        try:
            await b.close()
        except Exception:
            pass


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception:
        traceback.print_exc()
        sys.exit(1)
