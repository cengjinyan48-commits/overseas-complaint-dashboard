#!/usr/bin/env python3
"""
结案预警邮件通知脚本
- 扫描 Excel 中「结案状态=未结案」且「结案预警」日期已到的记录
- 按跟进人分组，发送 HTML 格式预警邮件
- 通过 GitHub Actions 每天定时执行
"""
import os
import sys
import smtplib
import logging
from datetime import datetime, timezone, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pandas as pd

# ============================================================
# Config — 优先读环境变量（GitHub Secrets），其次用默认值
# ============================================================
SMTP_SERVER   = os.getenv("SMTP_SERVER",   "mail.tcl.com")
SMTP_PORT     = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER     = os.getenv("SMTP_USER",     "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SENDER_NAME   = os.getenv("SENDER_NAME",   "海外客户服务部")
FROM_ADDR     = SMTP_USER  # 发件地址即登录账号

# 跟进人 → 邮箱映射
FOLLOWER_EMAILS = {
    "郑小平": "payne.zheng@tcl.com",
    "陈耀球": "kt_yorkchen@tcl.com",
    "曾靖衍": "jingyan.zeng@tcl.com",
    "黄忠成": "zhongcheng.huang@tcl.com",
    "方益勋": "kt_fangyx@tcl.com",
}

# Excel 路径（相对于项目根目录）
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "2026年海外客户投诉台账.xlsx")

# 北京时间
BJT = timezone(timedelta(hours=8))

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)


def load_data() -> pd.DataFrame:
    """加载并清洗客诉数据"""
    df = pd.read_excel(EXCEL_PATH, sheet_name="所有客诉", header=1)
    # 过滤无效行：编号必须是数字
    df["_num"] = pd.to_numeric(df["编号"], errors="coerce")
    df = df[df["_num"].notna()].copy()
    # 解析日期
    for col in ["投诉日期", "应结案日期", "结案预警"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def find_overdue_warnings(df: pd.DataFrame) -> pd.DataFrame:
    """找出「结案状态=未结案」且 结案预警日期 <= 今天 的记录"""
    today = pd.Timestamp(datetime.now(BJT).replace(tzinfo=None))
    mask = (
        (df["结案状态"] == "未结案") &
        df["结案预警"].notna() &
        (df["结案预警"] <= today)
    )
    overdue = df[mask].copy()
    overdue["超期天数"] = (today - overdue["结案预警"]).dt.days
    return overdue.sort_values(["跟进人", "超期天数"], ascending=[True, False])


def build_html_table(records: pd.DataFrame) -> str:
    """为某个跟进人生成 HTML 表格"""
    rows_html = ""
    for _, r in records.iterrows():
        warn_date = r["结案预警"].strftime("%Y-%m-%d") if pd.notna(r["结案预警"]) else "-"
        due_date  = r["应结案日期"].strftime("%Y-%m-%d") if pd.notna(r["应结案日期"]) else "-"
        desc = str(r["问题描述"])[:120] if pd.notna(r["问题描述"]) else ""
        rows_html += f"""
        <tr>
            <td style="padding:8px 10px;border-bottom:1px solid #eee;">{int(r['_num'])}</td>
            <td style="padding:8px 10px;border-bottom:1px solid #eee;">{r['国家或地区']}</td>
            <td style="padding:8px 10px;border-bottom:1px solid #eee;">{due_date}</td>
            <td style="padding:8px 10px;border-bottom:1px solid #eee;color:#e60012;font-weight:bold;">{warn_date}</td>
            <td style="padding:8px 10px;border-bottom:1px solid #eee;color:#e60012;font-weight:bold;">{int(r['超期天数'])} 天</td>
            <td style="padding:8px 10px;border-bottom:1px solid #eee;font-size:12px;color:#666;">{desc}</td>
        </tr>"""

    return f"""
    <table style="width:100%;border-collapse:collapse;font-family:'Microsoft YaHei',Arial,sans-serif;font-size:13px;">
        <thead>
            <tr style="background:#e60012;color:#fff;">
                <th style="padding:10px;text-align:left;">编号</th>
                <th style="padding:10px;text-align:left;">国家</th>
                <th style="padding:10px;text-align:left;">应结案日期</th>
                <th style="padding:10px;text-align:left;">结案预警</th>
                <th style="padding:10px;text-align:left;">超期</th>
                <th style="padding:10px;text-align:left;">问题描述</th>
            </tr>
        </thead>
        <tbody>{rows_html}</tbody>
    </table>"""


def send_email(to_addr: str, to_name: str, records: pd.DataFrame,
               total_count: int) -> bool:
    """发送预警邮件给指定跟进人"""
    today_str = datetime.now(BJT).strftime("%Y-%m-%d")

    subject = f"【客诉预警】{to_name}，您有 {len(records)} 条客诉超期未结案 - {today_str}"

    # ---- HTML Body ----
    html_body = f"""
    <div style="max-width:800px;margin:0 auto;font-family:'Microsoft YaHei',Arial,sans-serif;">
        <div style="background:#e60012;padding:20px;border-radius:8px 8px 0 0;">
            <h2 style="color:#fff;margin:0;">2026年海外客户投诉 — 结案预警提醒</h2>
        </div>
        <div style="background:#fff;padding:20px;border:1px solid #e0e0e0;border-top:none;">
            <p style="font-size:14px;color:#333;">
                <strong>{to_name}</strong>，您好：
            </p>
            <p style="font-size:14px;color:#333;">
                截至 <strong>{today_str}</strong>，您负责的海外客诉中有
                <span style="color:#e60012;font-weight:bold;font-size:18px;"> {len(records)} 条</span>
                已超过结案预警日期，请尽快处理。
            </p>
            {build_html_table(records)}
            <p style="margin-top:20px;font-size:12px;color:#999;">
                ※ 此邮件由「海外客户投诉数据看板」自动发送，请勿回复。<br>
                ※ 当前全量未结案客诉共 <strong>{total_count}</strong> 条。
            </p>
        </div>
        <div style="background:#f5f5f5;padding:12px 20px;border-radius:0 0 8px 8px;font-size:11px;color:#999;text-align:center;">
            海外客户服务部 · 数据看板 · {today_str}
        </div>
    </div>"""

    # ---- 构造邮件 ----
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{SENDER_NAME} <{FROM_ADDR}>"
    msg["To"] = f"{to_name} <{to_addr}>"
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    # ---- 发送 ----
    try:
        if SMTP_PORT == 465:
            server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=15)
        else:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=15)
            server.ehlo()
            if server.has_extn("STARTTLS"):
                server.starttls()
                server.ehlo()

        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(FROM_ADDR, to_addr, msg.as_string())
        server.quit()
        log.info(f"  ✅ 已发送给 {to_name} <{to_addr}> — {len(records)} 条预警")
        return True
    except Exception as e:
        log.error(f"  ❌ 发送给 {to_name} <{to_addr}> 失败: {e}")
        return False


def main():
    log.info("=" * 60)
    log.info("结案预警邮件通知 — 开始执行")
    log.info(f"SMTP: {SMTP_SERVER}:{SMTP_PORT}  |  发件: {SMTP_USER}")

    if not SMTP_USER or not SMTP_PASSWORD:
        log.error("SMTP 账号/密码未配置！请设置环境变量 SMTP_USER / SMTP_PASSWORD")
        sys.exit(1)

    # 1. 加载数据
    df = load_data()
    log.info(f"已加载 {len(df)} 条有效记录")

    # 2. 筛选超期未结案
    overdue = find_overdue_warnings(df)
    log.info(f"超期未结案: {len(overdue)} 条")

    if len(overdue) == 0:
        log.info("无超期未结案记录，无需发送预警。")
        return

    # 3. 统计全量未结案（含未超期的）
    total_unclosed = int((df["结案状态"] == "未结案").sum())
    log.info(f"全量未结案: {total_unclosed} 条")

    # 4. 按跟进人分组发送
    success = 0
    fail = 0
    for follower, group in overdue.groupby("跟进人"):
        email = FOLLOWER_EMAILS.get(follower)
        if not email:
            log.warning(f"  ⚠️  跟进人「{follower}」无邮箱映射，跳过 {len(group)} 条记录")
            fail += 1
            continue

        log.info(f"→ {follower} ({len(group)} 条预警)")
        if send_email(email, follower, group, total_unclosed):
            success += 1
        else:
            fail += 1

    log.info("=" * 60)
    log.info(f"执行完毕: 成功 {success} 封, 失败 {fail} 封")
    if fail > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
