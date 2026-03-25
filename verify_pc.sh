#!/bin/bash
# PC知识验证计划
# 执行方式: bash verify_pc.sh
# 验证维度: 1)产品知识完整性 2)二开介入点识别
# 输出: ~/AI/client-data/产品标准推荐/verify_*.md

OUT="/Users/limingheng/AI/client-data/产品标准推荐"

echo "=== PC知识验证批次 ==="
echo "PC_02 供应商延期交付..."
python3 scripts/search_features.py "供应商延期交付 订单分级 催交 冻结区" > /dev/null 2>&1

echo "PC_05 来料质量..."
python3 scripts/search_features.py "来料质量 不合格 分层检验 CAPA" > /dev/null 2>&1

echo "PC_06 对账付款..."
python3 scripts/search_features.py "对账付款 自动对账 三单匹配 发票" > /dev/null 2>&1

echo "PC_08 供应商准入..."
python3 scripts/search_features.py "供应商准入 资质管理 风险扫描" > /dev/null 2>&1

echo "PC_09 合同管理..."
python3 scripts/search_features.py "合同管理 合同签署 合同变更" > /dev/null 2>&1

echo "PC_11 供应商关系经营..."
python3 scripts/search_features.py "供应商关系 关键供应商 供应商分级" > /dev/null 2>&1

echo "PC_21 寻源方式..."
python3 scripts/search_features.py "寻源方式 招标 竞价 询价" > /dev/null 2>&1

echo "PC_25 收货验收..."
python3 scripts/search_features.py "收货验收 到场确认 送货单" > /dev/null 2>&1

echo "PC_27 供应商分类分级..."
python3 scripts/search_features.py "供应商分类分级 评分 评级" > /dev/null 2>&1

echo "PC_28 供应商绩效管理..."
python3 scripts/search_features.py "供应商绩效 考核 绩效评分" > /dev/null 2>&1

echo "PC_32 预测计划..."
python3 scripts/search_features.py "预测计划 滚动预测 配额 MRP" > /dev/null 2>&1

echo ""
echo "=== 今日生成的验证报告 ==="
ls -lt $OUT/需求推荐_20260325*.md 2>/dev/null | head -15
echo ""
echo "=== 下一步：读取报告，评估是否有重大知识遗漏和二开点 ==="
echo "查看最新报告: ls \$OUT/需求推荐_20260325*.md"
