#!/usr/bin/env python3
"""
从石墨文档（Shimo 企业私有部署）同步 Excel 数据到本地

使用 Playwright 浏览器自动化：
1. 打开石墨文档
2. 注入 cookies 和 localStorage（登录态）
3. 点击右上角「下载」按钮触发下载
4. 验证下载的文件是有效的 xlsx

用法:
    python sync_from_shimo.py
    SHIMO_AUTH="<base64>" python sync_from_shimo.py
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
    log.info("安装 Chromium 浏览器...")
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        capture_output=True, text=True, timeout=300,
        env={**os.environ, "PLAYWRIGHT_BROWSERS_PATH": "/tmp/ms-playwright"},
    )
    with open(marker, "w") as f:
        f.write("ok")
    log.info("Chromium 安装完成")


def _is_valid_xlsx(path):
    """检查文件是否为有效的 xlsx"""
    if not os.path.exists(path):
        return False
    if os.path.getsize(path) < 1000:
        return False
    with open(path, "rb") as f:
        return f.read(2) == b"PK"


def main():
    if not AUTH_B64:
        log.error("SHIMO_AUTH 环境变量未设置！")
        sys.exit(1)

    # 解码认证信息
    raw = AUTH_B64.strip().replace("\n", "").replace("\r", "").replace(" ", "")
    try:
        auth = json.loads(base64.b64decode(raw).decode("utf-8"))
    except Exception as e:
        log.error(f"SHIMO_AUTH 解码失败: {e}")
        sys.exit(1)

    cookies = auth.get("cookies", [])
    ls_data = auth.get("localStorage", {})
    log.info(f"已加载 {len(cookies)} 个 cookies, {len(ls_data)} 个 localStorage 条目")

    _ensure_chromium()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # 先访问域名设置 cookies
        page.goto("https://teamwork.getech.cn", timeout=15000, wait_until="domcontentloaded")
        context.add_cookies(cookies)

        # 打开文档
        page.goto(SHIMO_URL, timeout=30000, wait_until="domcontentloaded")
        time.sleep(5)

        # 注入 localStorage
        page.evaluate(
            "(items) => { for (let k in items) { try { localStorage.setItem(k, items[k]); } catch(e) {} } }",
            ls_data,
        )

        # 重新加载以应用 localStorage
        page.reload()
        time.sleep(8)

        current_url = page.url
        log.info(f"当前页面: {current_url}")

        # 检查是否被重定向到登录页
        if "login" in current_url.lower():
            log.warning("检测到登录页重定向，尝试清除遮罩层...")
            # 尝试清除可能的登录弹窗
            page.evaluate("""
                () => {
                    document.querySelectorAll('[class*="modal"], [class*="overlay"], [class*="mask"], [class*="shadow"]')
                        .forEach(el => el.remove());
                }
            """)
            time.sleep(2)
            # 尝试直接导航到编辑页
            page.goto(SHIMO_URL, timeout=15000, wait_until="domcontentloaded")
            time.sleep(5)
            current_url = page.url
            log.info(f"重试后页面: {current_url}")

        # 设置下载监听
        download_done = False

        def on_download(dl):
            nonlocal download_done
            tmp_path = OUTPUT_PATH + ".tmp"
            try:
                dl.save_as(tmp_path)
                if _is_valid_xlsx(tmp_path):
                    os.replace(tmp_path, OUTPUT_PATH)
                    download_done = True
                    log.info(f"下载成功: {dl.suggested_filename}")
                else:
                    log.warning(f"下载的文件无效: {os.path.getsize(tmp_path)} bytes")
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
            except Exception as e:
                log.warning(f"保存下载文件失败: {e}")

        page.on("download", on_download)

        # 策略1: 点击右上角「下载」按钮
        page.evaluate("""
            () => {
                for (let el of document.querySelectorAll('button, a, [role="button"], span, div')) {
                    const t = (el.textContent || '').trim();
                    if (t === '下载' || t === '导出') { el.click(); return 'clicked:' + t; }
                }
                return 'no button';
            }
        """)
        time.sleep(5)

        # 策略2: 如果有导出菜单，选 Excel
        if not download_done:
            page.evaluate("""
                () => {
                    for (let el of document.querySelectorAll('li, div, span, button')) {
                        const t = (el.textContent || '').trim();
                        if (t.includes('Excel') || t.includes('xlsx') || t.includes('导出为') || t.includes('另存为')) {
                            el.click(); return 'clicked:' + t;
                        }
                    }
                    return 'no export';
                }
            """)
            time.sleep(5)

        # 策略3: Ctrl+S
        if not download_done:
            log.info("尝试 Ctrl+S...")
            page.keyboard.press("Control+S")
            time.sleep(5)

        # 等待下载完成
        for _ in range(15):
            if download_done:
                break
            time.sleep(1)

        context.close()
        browser.close()

    # 验证结果
    if _is_valid_xlsx(OUTPUT_PATH):
        import pandas as pd
        df = pd.read_excel(OUTPUT_PATH, sheet_name="所有客诉", header=1, engine="openpyxl")
        df["_n"] = pd.to_numeric(df["编号"], errors="coerce")
        valid = df[df["_n"].notna() & df["分公司"].notna() & (df["分公司"].astype(str).str.strip() != "")]
        log.info(f"✅ 同步成功！{os.path.getsize(OUTPUT_PATH)} bytes, {len(valid)} 条记录")
    else:
        # 检查是否有备份可用
        if os.path.exists(OUTPUT_PATH):
            with open(OUTPUT_PATH, "rb") as f:
                head = f.read(100)
            log.error(f"下载失败，文件头: {head[:50]}")
            log.error("认证可能已过期，请重新运行 get_shimo_auth.py 获取新的 SHIMO_AUTH")
        else:
            log.error("下载失败，文件不存在")
        sys.exit(1)


if __name__ == "__main__":
    main()