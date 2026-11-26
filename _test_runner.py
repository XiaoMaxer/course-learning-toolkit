"""
FIF 脚本真机测试（非交互式）：
  1. 启动独立 Edge（测试profile + CDP端口）
  2. 如果有保存的登录态（_edge_state.json） → 自动登录
  3. 如果没有登录态 → 让用户手动登录一次，保存登录态
  4. 注入最新脚本，点击开始刷课，观察日志

用法：python _test_runner.py [--login] [--duration N]

  --login     强制弹出浏览器让你重新登录（会保存新的登录态）
  --duration  观察秒数（默认 60）
"""
import asyncio
import sys
import os
import json
import argparse
import time

from playwright.async_api import async_playwright, TimeoutError as PWTimeout

HERE = os.path.dirname(os.path.abspath(__file__))
USER_DATA = os.path.join(HERE, "_edge_test_profile")  # 独立测试profile（非默认，所以可以开CDP）
STATE_FILE = os.path.join(HERE, "_edge_state.json")     # 保存的登录态
SCRIPT_FILE = os.path.join(HERE, "course-learning-toolkit.private.user.js")
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
START_URL = "https://icourse.fifedu.com/istp-manage/work/home"
CDP_PORT = 9223


def has_valid_state():
    """检查是否有可用的保存登录态"""
    if not os.path.exists(STATE_FILE):
        return False
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            state = json.load(f)
        cookies = state.get("cookies", [])
        fif_cookies = [c for c in cookies if "fifedu" in c.get("domain", "")]
        return len(fif_cookies) > 0
    except Exception:
        return False


async def wait_for_login(page, timeout=180):
    """等用户手动登录，最多 timeout 秒"""
    print(f"[TEST] ⏳ 等待你手动登录 FIF（最多 {timeout} 秒）...")
    print(f"[TEST] 请在弹出的 Edge 窗口里登录 FIF 平台")
    try:
        await page.wait_for_url("**/work/home**", timeout=timeout * 1000)
        print("[TEST] ✅ 检测到登录成功！")
        return True
    except PWTimeout:
        print("[TEST] ❌ 登录超时")
        return False


async def run_test(duration=60, force_login=False):
    async with async_playwright() as p:
        # ===== 启动 Edge（测试profile + CDP端口） =====
        # 用 launch_persistent_context 才能用自定义 user_data_dir
        # 这个测试profile是独立的，不影响用户正在使用的Edge
        print(f"[TEST] 正在启动 Edge（测试profile: {USER_DATA}）...")
        context = await p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA,
            executable_path=EDGE_EXE,
            channel="msedge",
            headless=False,  # 必须显示，让用户能手动登录
            args=[
                f"--remote-debugging-port={CDP_PORT}",
                "--start-maximized",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-blink-features=AutomationControlled",
            ],
            viewport={"width": 1920, "height": 1080},
        )
        page = context.pages[0] if context.pages else await context.new_page()

        # ===== 检查登录态 =====
        has_state = has_valid_state()
        print(f"[TEST] 有保存的登录态文件: {has_state}")

        if force_login:
            print("[TEST] 强制登录模式：清除旧 cookie")
            try:
                await context.clear_cookies()
            except Exception as e:
                print(f"[WARN] 清 cookie 失败: {e}")

        # ===== 加载 FIF 页面 =====
        print(f"[TEST] 正在加载 {START_URL} ...")
        await page.goto(START_URL, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(3)

        cur_url = page.url
        print(f"[TEST] 当前 URL: {cur_url}")

        # 以实际跳转结果判断是否需要登录（持久化 profile 里可能已有 cookie）
        need_login = ("login" in cur_url.lower()) or ("sso" in cur_url.lower())

        if need_login:
            if await wait_for_login(page):
                # 登录成功，保存登录态
                state = await context.storage_state()
                with open(STATE_FILE, "w", encoding="utf-8") as f:
                    json.dump(state, f, ensure_ascii=False, indent=2)
                print(f"[TEST] ✅ 已保存登录态到 {STATE_FILE}（以后不用再手动登录了！）")
            else:
                print("[TEST] 登录失败，退出")
                await context.close()
                return
        else:
            print("[TEST] ✅ FIF 平台已登录！")

        # ===== 注入刷课脚本 =====
        if not os.path.exists(SCRIPT_FILE):
            print(f"[ERROR] 脚本文件不存在: {SCRIPT_FILE}")
            await context.close()
            return

        with open(SCRIPT_FILE, "r", encoding="utf-8") as f:
            script_content = f.read()

        # 用 add_init_script 让脚本在页面加载时自动注入
        # 然后 reload 页面让脚本生效
        print("[TEST] 注入刷课脚本...")
        await context.add_init_script(script_content)
        await page.reload(wait_until="domcontentloaded", timeout=30000)
        print(f"[TEST] 已注入脚本（{len(script_content)} 字符）")
        await asyncio.sleep(3)

        # ===== 检查刷课面板 =====
        panel_status = await page.evaluate("""() => {
            const panel = document.getElementById('fif-helper-panel');
            const toggleBtn = document.getElementById('fif-toggle-run') || document.querySelector('[data-cmd=toggle-run]');
            const clearBtn = document.getElementById('fif-clear-skip') || document.querySelector('[data-cmd=clear-skip]');
            return {
                panelVisible: !!panel && panel.offsetParent !== null,
                toggleBtnFound: !!toggleBtn,
                clearBtnFound: !!clearBtn,
                statusText: document.getElementById('fif-status-text')?.innerText || '无',
            };
        }""")
        print(f"[TEST] 面板状态: {json.dumps(panel_status, ensure_ascii=False, indent=2)}")

        # ===== 清空跳过缓存（干净状态） =====
        if panel_status.get("clearBtnFound"):
            cleared = await page.evaluate("""() => {
                const btn = document.getElementById('fif-clear-skip') || document.querySelector('[data-cmd=clear-skip]');
                if (btn) { btn.click(); return 'clicked'; }
                return 'not-found';
            }""")
            print(f"[TEST] 清空跳过缓存: {cleared}")
            await asyncio.sleep(1)

        # ===== 点击开始刷课 =====
        toggled = await page.evaluate("""() => {
            const btn = document.getElementById('fif-toggle-run') || document.querySelector('[data-cmd=toggle-run]');
            if (btn) {
                const oldText = btn.innerText;
                btn.click();
                return `clicked: ${oldText} → ${btn.innerText}`;
            }
            // fallback: 找任何含「开始刷课」的按钮
            const all = document.querySelectorAll('button, [role=button], .el-button');
            for (const b of all) {
                const txt = b.innerText;
                if (txt.includes('开始') && (txt.includes('刷课') || txt.includes('学习'))) {
                    b.click();
                    return `fallback-clicked: ${txt}`;
                }
            }
            return 'toggle-btn-not-found';
        }""")
        print(f"[TEST] 点击开始刷课: {toggled}")
        await asyncio.sleep(2)

        # ===== 观察 duration 秒 =====
        print(f"\n[TEST] 🔍 开始观察 {duration} 秒（Ctrl+C 可提前结束）...")
        print("=" * 50)
        
        try:
            for i in range(duration):
                await asyncio.sleep(1)
                # 每 10 秒抓一次状态
                if i % 10 == 0 and i > 0:
                    status = await page.evaluate("""() => {
                        return {
                            badge: document.getElementById('fif-mode-badge')?.innerText || '?',
                            running: document.getElementById('fif-status-text')?.innerText || '?',
                            url: location.href.split('?')[0],  // 去掉query更清晰
                            title: document.title.slice(0, 30),
                        };
                    }""")
                    print(f"  [{i:3d}s] {status['badge']} | {status['running']:20s} | {status['url']}")
        except KeyboardInterrupt:
            print("\n[TEST] 收到 Ctrl+C，提前结束观察")

        # ===== 抓取最终日志 =====
        logs = await page.evaluate("""() => {
            const box = document.getElementById('fif-log-box') || document.querySelector('[id*=log]');
            if (box) return box.innerText.split('\\n').slice(-60).join('\\n');
            return '(无日志框)';
        }""")
        status = await page.evaluate("""() => {
            return {
                badge: document.getElementById('fif-mode-badge')?.innerText || '?',
                running: document.getElementById('fif-status-text')?.innerText || '?',
                url: location.href,
                title: document.title,
            };
        }""")
        
        print("\n" + "=" * 50)
        print("[TEST] ====== 最近 60 条刷课日志 ======")
        print(logs if logs.strip() else "(日志为空)")
        print("[TEST] ====== 日志结束 ======")
        print(f"\n[TEST] 📊 最终状态:")
        print(json.dumps(status, ensure_ascii=False, indent=2))

        # 保持浏览器打开供调试
        print(f"\n[TEST] Edge 窗口保持打开（调试端口 {CDP_PORT}），按 Ctrl+C 关闭")
        print(f"[TEST] 你可以在这个 Edge 里手动操作，或者让我重新跑测试")
        try:
            while True:
                await asyncio.sleep(1)
        except KeyboardInterrupt:
            pass
        finally:
            await context.close()
            print("[TEST] Edge 已关闭，测试结束")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="FIF 刷课脚本真机测试（首次需要手动登录保存登录态）"
    )
    parser.add_argument("--login", action="store_true", help="强制重新登录（保存新登录态）")
    parser.add_argument("--duration", type=int, default=60, help="观察秒数（默认60）")
    args = parser.parse_args()

    asyncio.run(run_test(duration=args.duration, force_login=args.login))
