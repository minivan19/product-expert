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
**处理**（三步分析框架）：
1. **买了没？**（模块维度）→ 合同xlsx优先，其次客户主数据xlsx"购买模块"列 → 关键词匹配（不用LLM）
2. **实施了没？**（模块+功能维度）→ 蓝图方案文件夹（DOC/PPTX/PDF）→ 内容关键词匹配（不用LLM，蓝图书写标准）
3. **用了没？**（模块+功能维度）→ 运维工单xlsx（标题+模块+描述+根本原因+解决方案）→ LLM凝练（口语化）

三步结果合并后与产品功能清单（product_modules_hierarchy.json）对比，输出3×2网格分类 + 激活建议。

**输出**：3×2网格分类表 + 各分类激活/推荐建议

### 场景3：业务方案 → 实施路线图
**输入**：业务专家 Skill 输出的综合经营分析方案（Part6）
**处理**：读取方案内容 → 关联产品功能模块 → 规划实施顺序
**输出**：分阶段实施路线图 + 必要时的二开建议

## 知识库架构

### Qdrant Collection: product_knowledge
- **维度**：3072（text-embedding-3-large）
- **距离算法**：Cosine
- **Embedding API**：gptsapi.net（与 Mem0 一致）
- **LLM 分析**：DeepSeek（deepseek-chat via api.deepseek.com）
- **运维工单处理**：LLM 语义凝练（描述+解决方案+根本原因），仅取去年完整自然年数据
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

### 数据源优先级（场景2，三步分析）
| 步骤 | 数据源 | 格式 | 方法 | LLM |
|------|--------|------|------|-----|
| Step1 买了没 | 合同xlsx（优先）→ 主数据xlsx（备选） | .xlsx | 关键词匹配 | 否 |
| Step2 实施了没 | 蓝图方案文件夹 | .doc/.pptx/.pdf | 内容关键词匹配 | 否 |
| Step3 用了没 | 运维工单文件夹 | .xlsx | LLM语义凝练 | **是** |

**支持格式**：xlsx / PDF / DOC / DOCX（自动识别处理）

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
**用途**：场景2 — 三步缺口分析
**输入**：客户名称
**处理**：自动查找该客户的商务档案 + 蓝图方案 + 运维工单目录，执行三步分析
**输出**：Markdown 报告（3×2网格分类 + 激活建议）
**用法**：
```bash
python scripts/gap_analysis.py 诺斯贝尔
python scripts/gap_analysis.py 诺斯贝尔 --year 2025
python scripts/gap_analysis.py 诺斯贝尔 --output 缺口报告.md
```

### md2docx.py
**用途**：Markdown 转 Word 文档（独立工具）
**输入**：Markdown 文件
**输出**：Word 文档（.docx）
**用法**：
```bash
python scripts/md2docx.py --input report.md --output report.docx
python scripts/md2docx.py --input report.md --output report.docx --template template.docx
```
**注意**：支持无模板 fallback，不指定 `--template` 时生成纯样式 Word

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

1. **Step3 LLM语义凝练**：目前 Step3（用了没）仍用关键词匹配，LITE模式；完整LLM凝练待实现（需解决API调用效率问题）
2. **PDF流程图OCR**：Windows不支持PaddlePaddle，无法安装PaddleOCR；当前用文件名降级匹配；**待后续优化**：Linux环境部署OCR或接入云端OCR API
3. **合同数据粒度**：合同xlsx中的产品名为"甄云SRM V3.0"（产品级），无法精确到模块；实际以客户主数据"购买模块"列为主
4. **DOC读取**：Win32COM方式，约6个文件需要逐个处理，较大文件可能截断
5. **时间筛选**：仅支持"提单时间"列的去年自然年筛选，格式支持 yyyy-mm-dd、yyyy/mm/dd
