"""
真机测试：注入脚本 → 进入课程 → 自动刷课 → 观察 180 秒，同时抓控制台日志与面板状态。
用法：python _test_course.py [--duration 180]
"""
import asyncio
import argparse
import json
import os
import subprocess
import sys
import time
import urllib.request

from playwright.async_api import async_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE = os.path.join(HERE, "_edge_test_profile")
SCRIPT_FILE = os.path.join(HERE, "course-learning-toolkit.private.user.js")
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
CDP = "http://127.0.0.1:9223"
URL = "https://icourse.fifedu.com/istp-manage/work/home"


def cdp_version():
    try:
        with urllib.request.urlopen(CDP + "/json/version", timeout=2) as r:
            return json.load(r).get("Browser")
    except Exception:
        return None


def ensure_edge():
    if cdp_version():
        print("[TEST] CDP 已在线")
        return
    print("[TEST] 启动 Edge ...")
    subprocess.Popen([EDGE, "--remote-debugging-port=9223", f"--user-data-dir={PROFILE}",
                      "--start-maximized", "--no-first-run", URL], close_fds=True)
    for i in range(25):
        time.sleep(1)
        if cdp_version():
            print(f"[TEST] CDP 就绪 {i+1}s")
            return
    sys.exit("[TEST] CDP 超时")


async def main(duration):
    with open(SCRIPT_FILE, "r", encoding="utf-8") as f:
        script = f.read()

    console_msgs = []

    async with async_playwright() as p:
        b = await p.chromium.connect_over_cdp(CDP)
        ctx = b.contexts[0]
        page = [pg for pg in ctx.pages if pg.url.startswith("http")][0]

        def on_console(msg):
            try:
                console_msgs.append(f"[{msg.type}] {msg.text}")
            except Exception:
                pass
        page.on("console", on_console)
        page.on("pageerror", lambda e: console_msgs.append(f"[pageerror] {e}"))

        if "work/home" not in page.url:
            await page.goto(URL, wait_until="domcontentloaded", timeout=45000)
            await asyncio.sleep(4)

        # 注册 init script 并 reload，确保脚本在页面加载时运行
        await ctx.add_init_script(script)
        await page.reload(wait_until="domcontentloaded", timeout=45000)
        await asyncio.sleep(6)
        print(f"[TEST] 注入完成，URL={page.url}")

        panel = await page.evaluate("""() => ({
            hasPanel: !!document.getElementById('fif-helper-panel'),
            buttons: Array.from(document.querySelectorAll('button')).map(b => b.innerText.trim()).filter(Boolean).slice(0, 40),
        })""")
        print("[TEST] 面板存在:", panel["hasPanel"])
        print("[TEST] 按钮:", panel["buttons"])

        # 尝试找课程卡片上的「自动刷此课」按钮并点击
        clicked = await page.evaluate("""() => {
            const els = Array.from(document.querySelectorAll('button, [role=button], .el-button, a, span, div'));
            const b = els.find(x => /自动刷此课|自动刷此课程/.test(x.innerText || '') && (x.innerText || '').length < 20);
            if (b) { b.click(); return b.innerText.trim(); }
            return null;
        }""")
        print(f"[TEST] 点击自动刷此课: {clicked}")

        if not clicked:
            # 退而求其次：点击课程卡片本身
            clicked2 = await page.evaluate("""() => {
                const els = Array.from(document.querySelectorAll('div, a'));
                const card = els.find(e => /学习进度/.test(e.innerText || '') && (e.innerText || '').length < 200);
                if (card) { card.click(); return 'card-clicked: ' + card.innerText.replace(/\\s+/g,' ').slice(0,50); }
                return null;
            }""")
            print(f"[TEST] 点击课程卡片: {clicked2}")

        print(f"[TEST] 观察 {duration} 秒 ...")
        for i in range(duration):
            await asyncio.sleep(1)
            if i and i % 15 == 0:
                st = await page.evaluate("""() => {
                    const panel = document.getElementById('fif-helper-panel');
                    return {
                        url: location.href,
                        panel: panel ? panel.innerText.replace(/\\n+/g, ' | ') : '(无面板)',
                    };
                }""")
                print(f"\n[TEST] ==== {i}s ====")
                print(f"[TEST] URL: {st['url'][:130]}")
                print(f"[TEST] 面板: {st['panel'][:400]}")
                if console_msgs:
                    print(f"[TEST] 控制台最后3条: " + " || ".join(console_msgs[-3:])[:300])

        print("\n[TEST] ====== 控制台消息（最后40条） ======")
        for m in console_msgs[-40:]:
            print(m[:300])
        print("[TEST] ====== 结束 ======")

        final = await page.evaluate("""() => ({
            url: location.href,
            panel: document.getElementById('fif-helper-panel')?.innerText || '(无)',
        })""")
        print(f"[TEST] 最终 URL: {final['url']}")
        print("[TEST] 最终面板（尾部2000字）:")
        print(final["panel"][-2000:])

        try:
            await b.close()
        except Exception:
            pass


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--duration", type=int, default=180)
    a = ap.parse_args()
    ensure_edge()
    asyncio.run(main(a.duration))
