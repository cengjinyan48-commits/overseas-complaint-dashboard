#!/usr/bin/env python3
"""
从石墨文档同步数据 — 使用 API 直接下载（不依赖 Playwright 浏览器自动化）

用法:
    python sync_from_shimo.py                          # 使用 SHIMO_AUTH 环境变量
    SHIMO_AUTH="<base64>" python sync_from_shimo.py    # 直接指定
"""

import os, sys, json, base64, time, logging, io

import requests

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

SHIMO_URL = "https://teamwork.getech.cn/shimo-h5/shimo-edit/e1898e9f4b794a4786fcdfead749736c"
OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "2026年海外客户投诉台账.xlsx")
AUTH_B64 = os.getenv("SHIMO_AUTH", "")


def _is_valid_xlsx(path):
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

    # 提取关键认证信息
    cookies_dict = {}
    for c in auth.get("cookies", []):
        cookies_dict[c["name"]] = c["value"]

    ls = auth.get("localStorage", {})
    jwt_info_raw = ls.get("jwtInfo", "{}")
    try:
        jwt_info = json.loads(jwt_info_raw)
    except Exception:
        jwt_info = {}

    access_token = cookies_dict.get("accessToken", "")
    api_token = jwt_info.get("token", "")

    # 从 URL 提取 fileId
    file_id = SHIMO_URL.rstrip("/").split("/")[-1]

    log.info(f"accessToken: {'***' + access_token[-8:] if access_token else 'N/A'}")
    log.info(f"apiToken: {'***' + api_token[-8:] if api_token else 'N/A'}")
    log.info(f"fileId: {file_id}")

    # ── 方法1: 使用 Shimo API 导出 ──────────────────────────────
    session = requests.Session()

    # 设置 cookies
    for c in auth.get("cookies", []):
        session.cookies.set(c["name"], c["value"], domain=c.get("domain", ""), path=c.get("path", "/"))

    # 设置请求头
    token = api_token or access_token
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "Accept": "application/json, text/plain, */*",
        "Referer": SHIMO_URL,
        "Authorization": f"Bearer {token}",
    }

    # 尝试 API 导出
    export_url = f"https://teamwork.getech.cn/shimo-h5/api/v1/files/{file_id}/export"
    log.info(f"尝试 API 导出: POST {export_url}")

    try:
        resp = session.post(export_url, headers=headers, json={"type": "xlsx"}, timeout=60)
        log.info(f"API 响应: status={resp.status_code}, content-type={resp.headers.get('content-type', '')[:50]}")

        if resp.status_code == 200 and len(resp.content) > 1000:
            # 检查是否是 xlsx
            if resp.content[:2] == b"PK":
                with open(OUTPUT_PATH, "wb") as f:
                    f.write(resp.content)
                log.info(f"✅ API 导出成功！{len(resp.content)} bytes")
            elif b"<!DOCTYPE" in resp.content[:200] or b"<html" in resp.content[:200]:
                log.error(f"API 返回了 HTML 页面，认证可能已失效")
                log.error(f"响应片段: {resp.text[:500]}")
                sys.exit(1)
            else:
                log.warning(f"API 返回了未知格式，尝试保存: {len(resp.content)} bytes")
                with open(OUTPUT_PATH, "wb") as f:
                    f.write(resp.content)
        else:
            log.warning(f"API 导出失败: status={resp.status_code}, body={resp.text[:200]}")
    except Exception as e:
        log.warning(f"API 导出异常: {e}")

    # ── 方法2: 如果 API 失败，尝试 Playwright 浏览器自动化 ──────
    if not _is_valid_xlsx(OUTPUT_PATH):
        log.info("API 导出失败，回退到 Playwright 浏览器自动化...")
        _sync_with_playwright(auth)
    else:
        # 验证数据
        import pandas as pd
        df = pd.read_excel(OUTPUT_PATH, sheet_name="所有客诉", header=1, engine="openpyxl")
        df["_n"] = pd.to_numeric(df["编号"], errors="coerce")
        valid = df[df["_n"].notna() & df["分公司"].notna() & (df["分公司"].astype(str).str.strip() != "")]
        log.info(f"✅ 同步成功！{len(valid)} 条记录")


def _sync_with_playwright(auth):
    """Playwright 浏览器自动化备用方案"""
    from playwright.sync_api import sync_playwright

    cookies = auth.get("cookies", [])
    ls_data = auth.get("localStorage", {})

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
        log.info(f"页面: {page.url}")

        download_done = False

        def on_download(dl):
            nonlocal download_done
            tmp = OUTPUT_PATH + ".tmp"
            dl.save_as(tmp)
            if _is_valid_xlsx(tmp):
                os.replace(tmp, OUTPUT_PATH)
                download_done = True
                log.info(f"下载成功: {dl.suggested_filename}")

        page.on("download", on_download)

        # 点击下载按钮
        page.evaluate("""
            () => {
                for (let el of document.querySelectorAll('button, a, [role="button"], span')) {
                    if ((el.textContent||'').trim() === '下载') { el.click(); return 'ok'; }
                }
                return 'not found';
            }
        """)
        time.sleep(5)

        if not download_done:
            page.keyboard.press("Control+S")
            time.sleep(5)

        for _ in range(15):
            if download_done:
                break
            time.sleep(1)

        context.close()
        browser.close()

    if _is_valid_xlsx(OUTPUT_PATH):
        import pandas as pd
        df = pd.read_excel(OUTPUT_PATH, sheet_name="所有客诉", header=1, engine="openpyxl")
        df["_n"] = pd.to_numeric(df["编号"], errors="coerce")
        valid = df[df["_n"].notna() & df["分公司"].notna() & (df["分公司"].astype(str).str.strip() != "")]
        log.info(f"✅ Playwright 同步成功！{len(valid)} 条记录")
    else:
        log.error("Playwright 同步也失败了")
        sys.exit(1)


if __name__ == "__main__":
    main()