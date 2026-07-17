#!/bin/bash
# ============================================================================
# 一键同步：石墨文档 → 台账数据 → 天极标准化底表 → 推送到 GitHub
#
# 使用方法：
#   1. 在石墨文档中编辑完数据，保存
#   2. 终端运行: bash sync_and_push.sh
#   3. 打开 Streamlit 看板，点击「立即同步」
# ============================================================================
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "=========================================="
echo "  石墨文档 → GitHub 一键同步"
echo "=========================================="
echo ""

# Step 1: 从石墨文档下载最新数据
echo "📥 [1/4] 从石墨文档下载数据..."
python3 sync_from_shimo.py
echo ""

# Step 2: 转换为天极标准化数据底表
echo "🔄 [2/4] 转换为天极标准化数据底表..."
python3 convert_for_taiji.py
echo ""

# Step 3: 提交到 Git
echo "📦 [3/4] 提交到 Git..."
git add "2026年海外客户投诉台账.xlsx" "海外客诉台账_标准化数据.xlsx"
if git diff --staged --quiet; then
    echo "  数据无变化，跳过提交"
else
    git commit -m "data: 手动同步石墨文档 $(date +'%Y-%m-%d %H:%M')"
    echo "  已提交"
fi
echo ""

# Step 4: 推送到 GitHub
echo "🚀 [4/4] 推送到 GitHub..."
git push
echo ""

echo "=========================================="
echo "  ✅ 同步完成！"
echo "=========================================="
echo ""
echo "  现在打开 Streamlit 看板，点击「立即同步」即可"
echo "  https://customers-after-sales-overseas-complaint.streamlit.app"
echo ""