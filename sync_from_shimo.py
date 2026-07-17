#!/usr/bin/env python3
"""
从石墨文档（Shimo 企业私有部署）自动同步 Excel 数据

用法:
    python sync_from_shimo.py                          # 使用 SHIMO_AUTH 环境变量
    SHIMO_AUTH="<base64>" python sync_from_shimo.py    # 直接指定

输出: 2026年海外客户投诉台账.xlsx
"""

import os, sys, json, base64, time, logging

from playwright.sync_api import sync_playwright

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

SHIMO_URL = "https://teamwork.getech.cn/shimo-h5/shimo-edit/e1898e9f4b794a4786fcdfead749736c"
OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "2026年海外客户投诉台账.xlsx")
AUTH_B64 = os.getenv("SHIMO_AUTH", "")


def _ensure_chromium():
    import subprocess
    marker = "/tmp/.chromium_installed"
    if os.path.exists(marker):
        return
    log.info("正在安装 Chromium 浏览器...")
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        capture_output=True, text=True, timeout=300,
        env={**os.environ, "PLAYWRIGHT_BROWSERS_PATH": "/tmp/ms-playwright"},
    )
    with open(marker, "w") as f:
        f.write("ok")
    log.info("Chromium 安装完成")


def main():
    if not AUTH_B64:
        log.error("SHIMO_AUTH 环境变量未设置！请在 Streamlit Cloud Secrets 或 GitHub Secrets 中配置")
        sys.exit(1)

    _ensure_chromium()

    # 解码认证信息
    raw = AUTH_B64.strip()
    raw = "".join(raw.split())
    try:
        auth_bundle = json.loads(base64.b64decode(raw).decode("utf-8"))
    except Exception as e:
        log.error(f"SHIMO_AUTH 解码失败: {e}")
        sys.exit(1)

    cookies = auth_bundle.get("cookies", [])
    ls_data = auth_bundle.get("localStorage", {})
    log.info(f"已加载 {len(cookies)} 个 cookies, {len(ls_data)} 个 localStorage 条目")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        page.goto("https://teamwork.getech.cn", timeout=15000, wait_until="domcontentloaded")
        context.add_cookies(cookies)

        page.goto(SHIMO_URL, timeout=30000, wait_until="domcontentloaded")
        time.sleep(5)

        page.evaluate(
            "(items) => { for (let k in items) { try { localStorage.setItem(k, items[k]); } catch(e) {} } }",
            ls_data,
        )
        page.reload()
        time.sleep(8)

        log.info(f"当前页面: {page.url}")

        # ── 触发下载 ──────────────────────────────────────────
        download_done = False

        def on_download(dl):
            nonlocal download_done
            dl.save_as(OUTPUT_PATH)
            download_done = True
            log.info(f"下载成功: {dl.suggested_filename}")

        page.on("download", on_download)

        # 策略1: 点击右上角工具栏的「下载」按钮
        log.info("查找下载按钮...")
        clicked = page.evaluate("""
            () => {
                const btns = document.querySelectorAll('button, a, [role="button"], div[class*="toolbar"] span, div[class*="header"] span');
                for (let b of btns) {
                    const text = (b.textContent || '').trim();
                    if (text === '下载' || text === '导出' || text === '下载为') {
                        b.click();
                        return 'clicked: ' + text;
                    }
                }
                return 'no button found';
            }
        """)
        log.info(f"点击结果: {clicked}")
        time.sleep(5)

        # 策略2: 如果有导出菜单，点击 Excel 导出
        if not download_done:
            exported = page.evaluate("""
                () => {
                    const items = document.querySelectorAll('li, div[class*="menu-item"], div[class*="dropdown-item"], span, button');
                    for (let item of items) {
                        const text = (item.textContent || '').trim();
                        if (text.includes('Excel') || text.includes('xlsx') || text.includes('导出为') || text.includes('另存为')) {
                            item.click();
                            return 'clicked: ' + text;
                        }
                    }
                    return 'no export option';
                }
            """)
            log.info(f"导出选项: {exported}")
            time.sleep(5)

        # 策略3: 尝试 Ctrl+S
        if not download_done:
            log.info("尝试 Ctrl+S...")
            page.keyboard.press("Control+S")
            time.sleep(5)

        # 策略4: 尝试 API 导出
        if not download_done:
            log.info("尝试 API 导出...")
            page.evaluate("""
                () => {
                    const fileId = window.location.pathname.split('/').pop();
                    const exportUrl = '/shimo-h5/api/v1/files/' + fileId + '/export';
                    fetch(exportUrl, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ type: 'xlsx' })
                    }).then(r => r.blob()).then(blob => {
                        const url = URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = 'export.xlsx';
                        a.click();
                    }).catch(e => console.error(e));
                }
            """)
            time.sleep(5)

        # 等待下载完成
        for _ in range(10):
            if download_done:
                break
            time.sleep(1)

        # 验证结果：只要文件存在且大于 1KB 就算成功
        if os.path.exists(OUTPUT_PATH) and os.path.getsize(OUTPUT_PATH) > 1000:
            import pandas as pd
            df = pd.read_excel(OUTPUT_PATH, sheet_name="所有客诉", header=1)
            df["_n"] = pd.to_numeric(df["编号"], errors="coerce")
            valid = df[df["_n"].notna() & df["分公司"].notna() & (df["分公司"].astype(str).str.strip() != "")]
            log.info(f"✅ 同步成功！{os.path.getsize(OUTPUT_PATH)} bytes, {len(valid)} 条记录")
        else:
            page.screenshot(path="/tmp/shimo_debug.png")
            log.error("下载失败，调试截图已保存")
            sys.exit(1)

        context.close()
        browser.close()


if __name__ == "__main__":
    main()