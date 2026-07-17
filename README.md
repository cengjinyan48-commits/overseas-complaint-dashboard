# 2026年海外客户投诉数据看板

> TCL 海外客户投诉数据可视化分析平台，支持云端部署、自动同步、一键导出和结案预警。

**在线地址**: [https://overseas-complaint-dashboard.streamlit.app](https://overseas-complaint-dashboard.streamlit.app)

---

## 功能总览

### 📊 数据看板
- **KPI 指标卡**: 总投诉量 / 结案率 / 平均处理周期 / 未结案数 / 超期预警数
- **概览仪表盘**: 分公司分布、国家TOP15、月度趋势+结案率、结案状态环形图
- **故障分析**: 故障大类帕累托图、机型属性+大客户占比、故障×机型交叉矩阵
- **质量管理**: 跟进人工作量、完成周期分布、8D报告覆盖率、超期未结案预警
- **数据明细**: 全局搜索、多维度筛选、可排序数据表

### 📥 数据导出
- **CSV 一键导出 (ZIP)**: 14 个分析维度 CSV 文件，含明细数据、各维度统计、交叉表、超期预警
- **PPT 分析报告**: 8页 TCL 品牌风格报告，封面+摘要+区域/趋势/故障/质量分析+改进建议
- **每周周报 (Excel)**: 选择周次一键导出，含本周客诉(三模块)/汇总/情况通报/赔偿费用/材料费用 5个Sheet

### 🚨 结案预警
- **自动邮件通知**: 每天 9:00（北京时间）GitHub Actions 自动扫描超期未结案记录 → 按跟进人分组发送 HTML 邮件
- **看板一键发送**: 侧边栏实时显示预警统计，点击按钮即可批量发送
- **仅计算未结案**: 超期预警只计算「结案状态=未结案」的记录，排除「暂停」「结案」「关闭」

### 🔄 数据同步
- **石墨文档自动同步**: GitHub Actions 每小时自动从石墨在线文档拉取最新数据
- **手动上传**: 侧边栏拖拽上传 Excel 即时刷新
- **侧边栏一键触发**: 点击「立即同步」按钮手动触发同步任务

### 🏢 天极（FineBI）数据集成
- **自动格式转换**: 同步完成后自动生成 33 列标准化数据底表（含投诉年月、是否超期、结案标记等 8 个计算字段 + 数据字典）
- **侧边栏一键下载**: 点击「立即同步」后自动生成天极数据底表，侧边栏直接下载
- **GitHub Actions 自动转换**: 每小时同步任务中自动转换，仓库中始终有最新的标准化文件
- **上传即用**: 下载后直接上传天极系统，无需额外处理

### 🎨 品牌定制
- 内置 TCL 品牌 PPT 模板（奥运版），导出的 PPT 自动套用 TCL Logo、配色、字体
- 封面保留奥运五环+奥林匹克全球合作伙伴标识
- 内容页右上角自动带上品牌 Logo

---

## 技术栈

| 组件 | 用途 |
|------|------|
| [Streamlit](https://streamlit.io) | 数据看板 Web 框架 |
| [Plotly](https://plotly.com) | 交互式图表 |
| [Pandas](https://pandas.pydata.org) | 数据处理 |
| [python-pptx](https://python-pptx.readthedocs.io) | PPT 生成 |
| [Matplotlib](https://matplotlib.org) | PPT 内嵌图表渲染 |
| [Playwright](https://playwright.dev) | 金山文档浏览器自动化同步 |
| [GitHub Actions](https://github.com/features/actions) | 定时任务（邮件通知 + 数据同步） |
| [Streamlit Cloud](https://streamlit.io/cloud) | 免费云端部署 |

---

## 项目结构

```
├── streamlit_app.py                  # 主应用（Streamlit 看板）
├── send_warning_emails.py            # 结案预警邮件脚本
├── sync_from_shimo.py               # 石墨文档自动同步脚本
├── sync_from_kdocs.py               # 金山文档自动同步脚本（旧，已弃用）
├── convert_for_taiji.py              # 天极标准化数据底表转换脚本
├── generate_weekly_report.py         # 每周周报生成器
├── generate_dashboard.py             # 本地 HTML 看板生成器
├── requirements.txt                  # Python 依赖
├── assets/
│   └── TCL_template.pptx            # TCL 品牌 PPT 模板
├── .streamlit/
│   ├── config.toml                   # Streamlit 主题配置
│   └── secrets.toml                  # Streamlit Cloud Secrets (本地，不入git)
├── packages.txt                      # Chromium 系统依赖
├── requirements.txt                  # Python 依赖
├── .github/workflows/
│   ├── warning_email.yml             # 结案预警定时任务
│   └── kdocs_sync.yml               # 金山文档同步 + 天极数据转换
├── 导出模板/
│   └── TCL PPT标准模板（2025奥运版）.pptx  # 原始模板文件
├── 2026年海外客户投诉台账.xlsx        # 原始台账数据
└── 海外客诉台账_标准化数据.xlsx       # 天极标准化数据底表（自动生成）
```

---

## 本地运行

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动看板
streamlit run streamlit_app.py

# 3. 浏览器打开 http://localhost:8501
```

---

## 部署

项目部署于 [Streamlit Community Cloud](https://streamlit.io/cloud)，关联 GitHub 仓库自动部署。

每次 `git push` 后约 1-2 分钟自动上线。

---

## GitHub Actions 定时任务

| 任务 | 时间 | 说明 |
|------|------|------|
| 📧 结案预警邮件 | 每天 9:00 (BJT) | 扫描超期未结案 → 邮件通知跟进人 |
| 🔄 石墨文档同步 + 天极转换 | 每小时 (BJT) | 从石墨文档下载最新 Excel → 自动转换为天极标准化数据底表 → 提交 |

### GitHub Secrets 配置

| Secret | 说明 |
|--------|------|
| `SMTP_SERVER` / `SMTP_PORT` / `SMTP_USER` / `SMTP_PASSWORD` | 企业邮箱 SMTP 配置 |
| `KDOCS_COOKIES` | 石墨文档登录态 cookies（已弃用，保留兼容） |
| `SHIMO_AUTH` | 石墨文档完整认证信息（cookies + localStorage） |
| `GH_PAT` | GitHub Personal Access Token（用于自动提交推送） |

### Streamlit Cloud Secrets 配置

在 [Streamlit Cloud 控制台](https://share.streamlit.io) 的 App Settings → Secrets 中配置：

```toml
SHIMO_AUTH = "<base64-encoded shimo auth bundle>"
```

该值用于「立即同步」按钮的 Playwright 浏览器自动化登录金山文档。

---

## 数据更新方式

| 方式 | 操作 |
|------|------|
| 🟢 **石墨在线文档** | 编辑石墨文档 → 侧边栏点击「立即同步」→ 自动生成天极数据底表 → 下载上传天极 |
| 🟡 **本地上传** | 侧边栏拖拽 Excel 文件 → 即时刷新看板 |
| 🔵 **直接替换** | 替换仓库中 `2026年海外客户投诉台账.xlsx` → git push |
| 🏢 **天极看板更新** | 下载 `海外客诉台账_标准化数据.xlsx` → 上传到天极系统 → 刷新数据 |

---

## 跟进人邮箱映射

| 跟进人 | 邮箱 |
|--------|------|
| 郑小平 | payne.zheng@tcl.com |
| 陈耀球 | kt_yorkchen@tcl.com |
| 曾靖衍 | jingyan.zeng@tcl.com |
| 黄忠成 | zhongcheng.huang@tcl.com |
| 方益勋 | kt_fangyx@tcl.com |

---

© 2026 TCL 海外客户服务部
