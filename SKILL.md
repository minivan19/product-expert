---
name: product-expert
description: >-
  甄云SRM产品专家Skill。基于产品功能知识库（Qdrant向量库），提供功能推荐、缺口分析和实施规划服务。
  触发场景：
  (1) 用户描述客户需求或痛点，询问哪些产品功能可以解决
  (2) 用户提供客户名称，从合同/蓝图/工单提取已用功能后，对比产品功能清单做缺口分析
  (3) 基于业务专家Skill输出的综合方案，生成实施路线图（依赖business-expert Skill）
---

# 产品专家 Skill

基于产品功能知识库（Qdrant）和 LLM 分析能力，为客户成功经理提供功能推荐和实施规划服务。

## 核心能力

### 场景1：需求 → 功能推荐
**输入**：客户需求描述或痛点（自然语言）
**处理**：Qdrant 向量检索 → LLM 分析匹配
**输出**：功能推荐清单 + 推荐理由 + 来源依据

### 场景2：已用功能 → 缺口分析
**输入**：客户名称
**处理**：
1. 从客户档案（商务专家原始数据）提取合同中的模块清单
2. 从运维工单（运维专家原始数据）分析高频使用模块
3. 如有蓝图/用户手册，提取实施范围内的模块
4. 三者合并为"已用功能清单"
5. 与产品功能知识库做对比
**输出**：已用功能 vs 未覆盖需求对比表 + 优先级建议

### 场景3：业务方案 → 实施路线图
**输入**：业务专家 Skill 输出的综合经营分析方案（Part6）
**处理**：读取方案内容 → 关联产品功能模块 → 规划实施顺序
**输出**：分阶段实施路线图 + 必要时的二开建议

## 知识库架构

### Qdrant Collection: product_knowledge
- **维度**：1536（text-embedding-3-small）
- **距离算法**：Cosine
- **隔离**：独立于 Mem0 的 openclaw_memories collection

### 数据结构
每条记录包含：
```json
{
  "text": "功能描述内容（用于向量检索）",
  "metadata": {
    "source": "用户手册" | "迭代清单",
    "module": "供应商" | "寻源" | "订单" | "财务" | "大数据" | ...,
    "type": "标准功能" | "新增" | "修正",
    "version": "v2024" | "v2025Q1" | ...",
    "doc_name": "原始文档名称"
  }
}
```

### 数据来源
| 来源 | 路径 | 格式 | 更新频率 |
|------|------|------|----------|
| 产品功能手册 | `C:\Users\mingh\client-data\raw\产品功能\`（65份MD） | Markdown | 几乎不变 |
| 迭代功能清单 | `C:\Users\mingh\client-data\raw\产品功能\甄云SRM产品功能清单.xlsx` | Excel | 每年一次 |

### 检索优先级（场景2）
1. **蓝图**（如有）— 项目实施范围，最准确
2. **用户手册**（如有）— 客户定制化实施范围
3. **合同模块清单**（商务专家原始数据）— 采购模块
4. **运维工单**（运维专家原始数据）— 高频使用模块

## 脚本说明

### import_knowledge.py
**用途**：将产品手册导入 Qdrant
**输入**：65份 MD 文档 + 迭代清单 xlsx
**输出**：product_knowledge collection
**用法**：
```bash
python scripts/import_knowledge.py           # 全量导入
python scripts/import_knowledge.py --incremental  # 增量更新（对比文件更新时间）
```

### search_features.py
**用途**：场景1 — 需求到功能推荐
**输入**：客户需求描述（自然语言）
**输出**：功能推荐清单（含理由和来源）
**用法**：
```bash
python scripts/search_features.py "客户希望管控供应商准入和资质有效期"
```

### gap_analysis.py
**用途**：场景2 — 缺口分析
**输入**：客户名称
**处理**：自动查找该客户的商务档案 + 运维工单目录
**输出**：缺口分析报告（Markdown）
**用法**：
```bash
python scripts/gap_analysis.py 明阳电路
python scripts/gap_analysis.py 明阳电路 --output report.md
```

## 依赖其他 Skill 的数据格式

### 商务专家原始数据
- 路径模式：`C:\Users\mingh\client-data\raw\客户档案\<客户名>\商务信息\`
- 关键文件：合同附件（含模块清单）、报价单
- 提取字段：采购模块列表、合同金额

### 运维工单原始数据
- 路径模式：`C:\Users\mingh\client-data\raw\客户档案\<客户名>\运维工单\`
- 关键文件：工单明细 Excel
- 提取字段：问题模块分布、工单类型

### 蓝图/用户手册（如有）
- 路径模式：`C:\Users\mingh\client-data\raw\客户档案\<客户名>\`
- 关键文件：`蓝图*.pdf`、`*用户手册*.md`
- 提取字段：实施模块清单、客户定制化功能

## 与其他 Skill 的协作

```
business-expert (场景3输入)
       ↓ Part6综合经营分析方案
product-expert (场景3)
       ↓ 实施路线图
[最终输出给客户]
```

场景3依赖 business-expert Skill 的输出，需要等业务专家完成后方可使用。

## 已知约束

1. **场景2 的已用功能提取**：部分客户可能缺少蓝图和用户手册，此时仅依赖合同模块 + 运维工单，完整性受限
2. **迭代清单更新**：每年才同步一次，知识库可能滞后于实际产品迭代
3. **embedding 模型**：使用 gptsapi.net 的 text-embedding-3-small，与 Mem0 保持一致
