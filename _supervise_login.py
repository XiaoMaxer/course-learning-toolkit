"""清理登录态并打开 FIF 登录页，输出表单结构"""
import asyncio
import json
import os
import subprocess
import sys
import time
import urllib.request

from playwright.async_api import async_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE = os.path.join(HERE, "_edge_test_profile")
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
CDP = "http://127.0.0.1:9223"
LOGIN = "https://www.fifedu.com/iplat/fifLogin/index.html?service=https%3A%2F%2Ficourse.fifedu.com%2Fistp-manage%2Fwork%2Fhome"


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
    for _ in range(25):
        time.sleep(1)
        if cdp_v():
            return
    sys.exit("CDP 超时")


async def main():
    async with async_playwright() as p:
        b = await p.chromium.connect_over_cdp(CDP)
        ctx = b.contexts[0]
        # 清 cookie 和 localStorage，确保是干净的新账号会话
        await ctx.clear_cookies()
        page = [pg for pg in ctx.pages if pg.url.startswith("http")][0]
        await page.goto(LOGIN, wait_until="domcontentloaded", timeout=45000)
        await asyncio.sleep(4)
        try:
            await page.evaluate("() => { localStorage.clear(); sessionStorage.clear(); }")
        except Exception:
            pass
        print("URL:", page.url)

        for f in page.frames:
            try:
                info = await f.evaluate("""() => ({
                    url: location.href,
                    inputs: Array.from(document.querySelectorAll('input')).map(i => ({
                        type: i.type, name: i.name, id: i.id,
                        ph: i.placeholder, cls: (i.className || '').slice(0, 50), visible: !!i.offsetParent,
                    })),
                    btns: Array.from(document.querySelectorAll('button, .btn, [type=submit], a')).map(x => (x.innerText || '').trim()).filter(Boolean).slice(0, 25),
                    text: (document.body?.innerText || '').replace(/\\s+/g, ' ').slice(0, 300),
                })""")
                print("\n--- frame:", f.url[:100])
                print(json.dumps(info, ensure_ascii=False, indent=1)[:2500])
            except Exception as e:
                print("frame 读取失败", f.url[:80], e)
        await b.close()


ensure_edge()
asyncio.run(main())
