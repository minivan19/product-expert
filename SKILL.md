---
name: product-expert
description: >-
  甄云SRM产品专家Skill。基于产品功能知识库（Qdrant向量库），提供功能推荐、缺口分析和实施规划服务。
  触发场景：
  (1) 用户描述客户需求或痛点，询问哪些产品功能可以解决
  (2) 用户提供客户名称，从合同/蓝图/工单提取已用功能后，对比产品功能清单做缺口分析
  (3) 基于业务专家Skill输出的综合方案，生成实施路线图（依赖business-expert Skill）
  (4) 管理产品方案卡资产（从蓝图提取、场景三生成入库、定期Review）
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
**输入**：business-expert 输出的结构化JSON（推荐处理方式，不含产品映射）
**处理**：
1. 接收业务专家JSON
2. 查产品方案卡库（PC_XX）——是否有匹配？
   - 有 → 复用现有卡片，CSM确认适用性
   - 无 → Qdrant检索处理方式对应的产品功能
3. 结合客户现状（场景二结果）→ 判断哪些模块需要新实施
4. 排实施顺序（基础数据 → 核心流程 → 高级功能）
5. 询问CSM是否将此方案入库为产品方案卡
**输出**：分阶段实施路线图 + 必要时的二开建议

### 场景4：产品方案卡管理
产品方案卡是「业务场景→产品模块组合→实施步骤」的标准化资产，来源：蓝图提取 / 场景三生成 / 手工创建。

**场景4a：蓝图方案 → 提取产品方案卡**
```
python3 scripts/extract_pc_from_blueprint.py --customer "诺斯贝尔" --auto
```
- 自动找客户最新蓝图方案文件
- LLM分析文本 → 凝练为产品方案卡（一对多）
- 每张卡片供CSM确认后保存

**场景4b：场景三方案 → 询问入库**
- 场景三输出实施路线图后，自动询问CSM是否保存为产品方案卡
- 保存后状态为"草稿"，需后续完善

**场景4c：卡片库Review**
```
python3 scripts/extract_pc_from_blueprint.py --review
```
- 扫描产品模块覆盖空白
- 提示建议补充的高频场景

**产品方案卡结构**：`framework/product_solution_card_schema.json`
**卡片索引**：`framework/product_card_index.json`

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
| 产品功能手册 | `/Users/limingheng/AI/client-data\raw\产品功能\`（65份MD） | Markdown | 几乎不变 |
| 迭代功能清单 | `/Users/limingheng/AI/client-data\raw\产品功能\甄云SRM产品功能清单.xlsx` | Excel | 每年一次 |

### 数据源优先级（场景2，三步分析）
| 步骤 | 数据源 | 格式 | 方法 | LLM |
|------|--------|------|------|-----|
| Step1 买了没 | 合同xlsx（优先）→ 主数据xlsx（备选） | .xlsx | 关键词匹配 | 否 |
| Step2 实施了没 | 蓝图方案文件夹 | .doc/.pptx/.pdf | 内容关键词匹配 | 否 |
| Step3 用了没 | 运维工单文件夹 | .xlsx | DeepSeek LLM并行提取（5并发） | **是** |

**支持格式**：xlsx / PDF / DOC / DOCX（自动识别处理）

**Step3 术语映射系统**：
- LLM 提取工单中的业务术语（如"采购订单"、"送货单"、"供应商准入"）
- 术语通过 `term_map.py` 映射到 (模块, 功能) 对
- 内置 `BUILTIN_TERM_MAP`（50+ 条核心映射）+ `term_feedback.json`（用户确认的扩展映射）
- 新术语自动进入待确认队列，通过交互完成映射闭环
- API 配置从 `~/.openclaw/openclaw.json` 自动读取 DeepSeek key

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
**默认输出路径**：`/Users/limingheng/AI/client-data/产品标准推荐/需求推荐_{时间戳}.md`
**用法**：
```bash
python scripts/search_features.py "客户希望管控供应商准入和资质有效期"
python scripts/search_features.py "客户希望管控供应商准入和资质有效期" --output /path/to/report.md
```

### gap_analysis.py
**用途**：场景2 — 三步缺口分析
**输入**：客户名称
**处理**：自动查找该客户的商务档案 + 蓝图方案 + 运维工单目录，执行三步分析
**输出**：Markdown 报告（3×2网格分类 + 激活建议）
**默认输出路径**：`/Users/limingheng/AI/client-data/{客户名}/缺口分析_{客户名}.md`
**用法**：
```bash
python scripts/gap_analysis.py 诺斯贝尔
python scripts/gap_analysis.py 诺斯贝尔 --year 2025
python scripts/gap_analysis.py 诺斯贝尔 --output 缺口报告.md
```

### term_map.py（内部模块）
**用途**：Step3 术语提取 + 映射管理
**功能**：
- `BUILTIN_TERM_MAP`：50+ 条核心业务术语 → (模块, 功能) 映射
- `term_feedback.json`：用户确认的扩展映射，持久化存储
- `extract_terms_via_llm()`：DeepSeek LLM 并行提取（5并发，13batch约30秒）
- `analyze_workorders()`：主分析流程，交互确认新术语
**API配置**：从 `~/.openclaw/openclaw.json` 自动读取 DeepSeek 配置

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
- 路径模式：`/Users/limingheng/AI/client-data\raw\客户档案\<客户名>\商务信息\`
- 关键文件：合同附件（含模块清单）、报价单
- 提取字段：采购模块列表、合同金额

### 运维工单原始数据
- 路径模式：`/Users/limingheng/AI/client-data\raw\客户档案\<客户名>\运维工单\`
- 关键文件：工单明细 Excel
- 提取字段：问题模块分布、工单类型

### 蓝图/用户手册（如有）
- 路径模式：`/Users/limingheng/AI/client-data\raw\客户档案\<客户名>\`
- 关键文件：`蓝图*.pdf`、`*用户手册*.md`
- 提取字段：实施模块清单、客户定制化功能

## 与 business-expert 的协作

```
business-expert 场景一/二
       ↓ 纯业务判断（不含产品映射）
       JSON: {推荐处理方式, 业务域, 约束条件, 方案结构}
              ↓
product-expert 场景三
       ↓ 查产品方案卡库 / Qdrant检索
       ↓ 结合客户现状（场景二）
       实施路线图 + 二开建议
              ↓
       询问CSM: 是否入库为产品方案卡？
              ↓
product-expert 场景四（卡片管理）
```

## 已知约束

1. **PDF流程图OCR**：Windows不支持PaddlePaddle，无法安装PaddleOCR；当前用文件名降级匹配；**待后续优化**：Linux环境部署OCR或接入云端OCR API
2. **合同数据粒度**：合同xlsx中的产品名为"甄云SRM V3.0"（产品级），无法精确到模块；实际以客户主数据"购买模块"列为主
3. **DOC读取**：Win32COM方式，约6个文件需要逐个处理，较大文件可能截断
4. **时间筛选**：仅支持"提单时间"列的去年自然年筛选，格式支持 yyyy-mm-dd、yyyy/mm/dd
5. **术语映射系统**：term_feedback.json 会随分析自动积累新术语；待确认术语通过交互完成映射闭环
