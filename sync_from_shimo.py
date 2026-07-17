#!/usr/bin/env python3
"""
从石墨文档（Shimo）自动同步 Excel 数据
由 GitHub Actions 定时执行或手动触发

与金山文档同步脚本的区别：
- 石墨文档域名为 teamwork.getech.cn（企业私有部署）
- 下载按钮在右上角工具栏
- 导出方式为点击"下载"按钮 → 选择 Excel 格式
"""

import os
import sys
import json
import base64
import time
import logging

from playwright.sync_api import sync_playwright

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# ── 配置 ──────────────────────────────────────────────────────────
SHIMO_URL = "https://teamwork.getech.cn/shimo-h5/shimo-edit/e1898e9f4b794a4786fcdfead749736c"
SHIMO_DOMAIN = "https://teamwork.getech.cn"
OUTPUT_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "2026年海外客户投诉台账.xlsx",
)
AUTH_B64 = os.getenv("SHIMO_AUTH", os.getenv("KDOCS_AUTH", ""))


def _ensure_chromium():
    """确保 Chromium 浏览器已安装"""
    import subprocess

    marker = "/tmp/.chromium_installed"
    if os.path.exists(marker):
        return
    log.info("正在安装 Chromium 浏览器（首次约30秒）...")
    result = subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        capture_output=True, text=True, timeout=300,
        env={**os.environ, "PLAYWRIGHT_BROWSERS_PATH": "/tmp/ms-playwright"},
    )
    if result.returncode == 0:
        with open(marker, "w") as f:
            f.write("ok")
        log.info("Chromium 安装完成")
    else:
        log.warning("Chromium 安装失败: " + result.stderr[-200:])


def main():
    if not AUTH_B64:
        log.error("SHIMO_AUTH 环境变量未设置！")
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

        # ── 第一步：先访问域名，设置 cookies ──────────────────────
        page.goto(SHIMO_DOMAIN, timeout=15000, wait_until="domcontentloaded")
        context.add_cookies(cookies)

        # ── 第二步：打开石墨文档 ──────────────────────────────────
        page.goto(SHIMO_URL, timeout=30000, wait_until="domcontentloaded")
        time.sleep(5)

        # 恢复 localStorage
        page.evaluate(
            "(items) => { for (let k in items) { try { localStorage.setItem(k, items[k]); } catch(e) {} } }",
            ls_data,
        )
        page.reload()
        time.sleep(8)

        # 确认登录状态
        has_login = page.evaluate(
            '() => !!document.querySelector("[class*=login]") || !!document.querySelector("[class*=auth]")'
        )
        if has_login:
            # 石墨文档可能显示登录页面，检查是否真的需要登录
            url = page.url
            if "login" in url.lower() or "auth" in url.lower():
                log.error("登录态已失效，请重新获取 SHIMO_AUTH！")
                page.screenshot(path="/tmp/shimo_debug.png")
                sys.exit(1)
        log.info("石墨文档已打开")

        # ── 第三步：触发下载 ──────────────────────────────────────
        download_done = False

        def on_download(dl):
            nonlocal download_done
            dl.save_as(OUTPUT_PATH)
            download_done = True
            log.info(f"下载成功: {dl.suggested_filename}")

        page.on("download", on_download)

        # 策略1：点击右上角「下载」按钮
        # 石墨文档的下载按钮常见选择器
        download_selectors = [
            'button:has-text("下载")',
            '[class*="download"]',
            '[class*="Download"]',
            '[class*="toolbar"] button:has-text("下载")',
            '[class*="header"] button:has-text("下载")',
            'text="下载"',
            'button[title*="下载"]',
            'button[title*="download"]',
            # 石墨文档文件菜单 → 导出
            '[class*="menu"] text="导出"',
            'text="导出为Excel"',
            'text="导出为 Excel"',
        ]

        for selector in download_selectors:
            if download_done:
                break
            try:
                btn = page.locator(selector).first
                if btn and btn.is_visible(timeout=2000):
                    log.info(f"找到下载按钮: {selector}")
                    btn.click(force=True, timeout=5000)
                    time.sleep(3)
                    # 如果有导出选项，选 Excel
                    excel_btn = page.locator('text="Excel"').first
                    if excel_btn:
                        try:
                            excel_btn.click(timeout=3000)
                            time.sleep(3)
                        except Exception:
                            pass
            except Exception:
                continue

        # 策略2：尝试文件菜单
        if not download_done:
            log.info("尝试文件菜单...")
            file_menu_selectors = [
                'text="文件"',
                '[class*="file"]',
                'button:has-text("文件")',
            ]
            for sel in file_menu_selectors:
                if download_done:
                    break
                try:
                    btn = page.locator(sel).first
                    if btn and btn.is_visible(timeout=2000):
                        btn.click(timeout=3000)
                        time.sleep(2)
                        # 在菜单中找导出/下载
                        for action in ["导出", "下载为", "另存为"]:
                            try:
                                item = page.locator(f'text="{action}"').first
                                if item:
                                    item.click(timeout=3000)
                                    time.sleep(3)
                                    # 再找 Excel 选项
                                    excel = page.locator('text="Excel"').first
                                    if excel:
                                        try:
                                            excel.click(timeout=3000)
                                            time.sleep(3)
                                        except Exception:
                                            pass
                                    break
                            except Exception:
                                continue
                except Exception:
                    continue

        # 策略3：Ctrl+S（石墨文档的保存快捷键）
        if not download_done:
            log.info("尝试 Ctrl+S...")
            page.keyboard.press("Control+S")
            time.sleep(5)

        # 策略4：尝试通过 JS 触发导出 API
        if not download_done:
            log.info("尝试 JS 触发导出...")
            page.evaluate(
                "() => {"
                "  let btns = document.querySelectorAll('button');"
                "  for (let b of btns) {"
                "    if (b.textContent.includes('下载') || b.textContent.includes('导出')) {"
                "      b.click();"
                "      return;"
                "    }"
                "  }"
                "}"
            )
            time.sleep(5)

        # ── 第四步：验证结果 ──────────────────────────────────────
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
            page.screenshot(path="/tmp/shimo_debug.png")
            log.error("下载未触发，调试截图已保存到 /tmp/shimo_debug.png")
            sys.exit(1)

        context.close()
        browser.close()


if __name__ == "__main__":
    main()