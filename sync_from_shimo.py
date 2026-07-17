#!/usr/bin/env python3
"""
从石墨文档（Shimo 企业私有部署）自动同步 Excel 数据
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


def _is_valid_xlsx(path):
    """检查文件是否为有效的 xlsx（PK 开头）"""
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

    _ensure_chromium()

    try:
        raw = AUTH_B64.strip()
        raw = "".join(raw.split())
        auth_bundle = json.loads(base64.b64decode(raw).decode("utf-8"))
    except Exception as e:
        log.error(f"SHIMO_AUTH 解码失败: {e}")
        sys.exit(1)

    cookies = auth_bundle.get("cookies", [])
    ls_data = auth_bundle.get("localStorage", {})
    log.info(f"已加载 {len(cookies)} 个 cookies, {len(ls_data)} 个 localStorage 条目")

    # 备份当前有效文件，防止下载失败破坏数据
    backup_path = OUTPUT_PATH + ".backup"
    if _is_valid_xlsx(OUTPUT_PATH):
        import shutil
        shutil.copy2(OUTPUT_PATH, backup_path)

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

        download_done = False
        def on_download(dl):
            nonlocal download_done
            # 先保存到临时文件，验证后再替换
            tmp_path = OUTPUT_PATH + ".tmp"
            dl.save_as(tmp_path)
            if _is_valid_xlsx(tmp_path):
                os.replace(tmp_path, OUTPUT_PATH)
                download_done = True
                log.info(f"下载成功: {dl.suggested_filename}")
            else:
                # 下载的是 HTML 或其他非 xlsx 内容
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
                log.warning("下载的文件不是有效的 xlsx")
        page.on("download", on_download)

        # 策略1: 点击下载按钮
        page.evaluate("""
            () => {
                const all = document.querySelectorAll('button, a, [role="button"], span');
                for (let el of all) {
                    const t = (el.textContent || '').trim();
                    if (t === '下载' || t === '导出') { el.click(); return 'ok:' + t; }
                }
                return 'not found';
            }
        """)
        time.sleep(5)

        # 策略2: 导出菜单
        if not download_done:
            page.evaluate("""
                () => {
                    const all = document.querySelectorAll('li, div, span, button');
                    for (let el of all) {
                        const t = (el.textContent || '').trim();
                        if (t.includes('Excel') || t.includes('导出为') || t.includes('另存为')) {
                            el.click(); return 'ok:' + t;
                        }
                    }
                    return 'not found';
                }
            """)
            time.sleep(5)

        # 策略3: Ctrl+S
        if not download_done:
            page.keyboard.press("Control+S")
            time.sleep(5)

        # 策略4: API 导出
        if not download_done:
            page.evaluate("""
                () => {
                    const fid = window.location.pathname.split('/').pop();
                    fetch('/shimo-h5/api/v1/files/' + fid + '/export', {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({type: 'xlsx'})
                    }).then(r => r.blob()).then(blob => {
                        const a = document.createElement('a');
                        a.href = URL.createObjectURL(blob);
                        a.download = 'export.xlsx'; a.click();
                    }).catch(e => console.error(e));
                }
            """)
            time.sleep(5)

        # 等待下载完成
        for _ in range(15):
            if download_done:
                break
            time.sleep(1)

        context.close()
        browser.close()

    # 验证下载结果
    if _is_valid_xlsx(OUTPUT_PATH):
        import pandas as pd
        df = pd.read_excel(OUTPUT_PATH, sheet_name="所有客诉", header=1, engine='openpyxl')
        df["_n"] = pd.to_numeric(df["编号"], errors="coerce")
        valid = df[df["_n"].notna() & df["分公司"].notna() & (df["分公司"].astype(str).str.strip() != "")]
        log.info(f"✅ 同步成功！{os.path.getsize(OUTPUT_PATH)} bytes, {len(valid)} 条记录")
        # 清理备份
        if os.path.exists(backup_path):
            os.remove(backup_path)
    else:
        # 下载失败，恢复备份
        if os.path.exists(backup_path) and _is_valid_xlsx(backup_path):
            import shutil
            shutil.copy2(backup_path, OUTPUT_PATH)
            log.warning("下载失败，已恢复旧数据")
        sys.exit(1)


if __name__ == "__main__":
    main()