#!/usr/bin/env python3
"""
从金山文档自动同步 Excel 数据
由 GitHub Actions 定时执行或手动触发
"""
import os, sys, json, base64, time, io, logging
from playwright.sync_api import sync_playwright

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

KDOCS_URL = "https://www.kdocs.cn/l/cjP6zkIRj17V"
OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "2026年海外客户投诉台账.xlsx")
AUTH_B64 = os.getenv("KDOCS_AUTH", "")


def _ensure_chromium():
    """确保 Chromium 浏览器已安装"""
    import subprocess
    marker = os.path.join(os.path.dirname(__file__), ".chromium_installed")
    if os.path.exists(marker):
        return
    log.info("正在安装 Chromium 浏览器（首次约30秒）...")
    result = subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        capture_output=True, text=True, timeout=300,
    )
    if result.returncode == 0:
        with open(marker, "w") as f:
            f.write("ok")
        log.info("Chromium 安装完成")
    else:
        log.warning("Chromium 安装失败: " + result.stderr[-200:])


def main():
    if not AUTH_B64:
        log.error("KDOCS_AUTH 环境变量未设置！")
        sys.exit(1)

    _ensure_chromium()

    # 解码认证信息
    auth_bundle = json.loads(base64.b64decode(AUTH_B64).decode())
    cookies = auth_bundle.get("cookies", [])
    ls_data = auth_bundle.get("localStorage", {})
    log.info(f"已加载 {len(cookies)} 个 cookies, {len(ls_data)} 个 localStorage 条目")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # 先访问域名设置 cookies
        page.goto("https://www.kdocs.cn", timeout=15000, wait_until="domcontentloaded")
        context.add_cookies(cookies)

        # 打开文档并恢复 localStorage
        page.goto(KDOCS_URL, timeout=30000, wait_until="domcontentloaded")
        time.sleep(5)
        page.evaluate(
            "(items) => { for (let k in items) { try { localStorage.setItem(k, items[k]); } catch(e) {} } }",
            ls_data,
        )
        page.reload()
        time.sleep(8)

        # 确认无登录弹窗
        has_modal = page.evaluate(
            '() => !!document.querySelector("[class*=login-modal]")'
        )
        if has_modal:
            log.error("登录态已失效，请重新获取 KDOCS_AUTH！")
            sys.exit(1)
        log.info("登录态有效")

        # 关闭弹窗遮罩
        page.evaluate(
            '() => document.querySelectorAll(".shadow").forEach(el => el.remove())'
        )
        time.sleep(0.5)

        # 下载文件
        download_done = False

        def on_download(dl):
            nonlocal download_done
            dl.save_as(OUTPUT_PATH)
            download_done = True
            log.info(f"下载成功: {dl.suggested_filename}")

        page.on("download", on_download)

        # 策略1: 更多菜单 → 另存
        menu_btn = page.query_selector('[class*="app-header-more-btn"]')
        if menu_btn:
            log.info("找到更多菜单按钮")
            menu_btn.click(force=True, timeout=8000)
            time.sleep(2)
        else:
            log.warning("未找到更多菜单按钮，尝试直接点击...")

        # 尝试多种方式触发下载
        for action in ["另存", "下载", "导出为"]:
            if download_done:
                break
            try:
                btn = page.locator(f"text={action}").first
                if btn:
                    log.info(f"点击 '{action}'...")
                    btn.click(force=True, timeout=5000)
                    time.sleep(10)
            except Exception as e:
                log.warning(f"点击 '{action}' 失败: {e}")

        # 策略2: 如果还没触发，尝试快捷键 Ctrl+S
        if not download_done:
            log.info("尝试 Ctrl+S 快捷键...")
            page.keyboard.press("Control+S")
            time.sleep(10)

        # 策略3: 尝试页面 JS 触发下载
        if not download_done:
            log.info("尝试 JS 触发下载...")
            page.evaluate(
                "() => { let a = document.createElement('a'); "
                "a.href = window.location.href.replace('/l/', '/api/v3/office/file/') + '/export/xlsx'; "
                "a.click(); }"
            )
            time.sleep(5)

        if download_done and os.path.exists(OUTPUT_PATH):
            import pandas as pd

            df = pd.read_excel(OUTPUT_PATH, sheet_name="所有客诉", header=1)
            df["_n"] = pd.to_numeric(df["编号"], errors="coerce")
            valid = df[
                df["_n"].notna()
                & df["分公司"].notna()
                & (df["分公司"].astype(str).str.strip() != "")
            ]
            log.info(
                f"✅ 同步成功！{os.path.getsize(OUTPUT_PATH)} bytes, {len(valid)} 条记录"
            )
        else:
            # 截图保存用于调试
            page.screenshot(path="/tmp/kdocs_debug.png")
            log.error("下载未触发，调试截图已保存")
            sys.exit(1)

        context.close()
        browser.close()


if __name__ == "__main__":
    main()
