#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
场景2：客户功能缺口分析（支持DOC/PDF/xlsx多格式）
用法：
  python scripts/gap_analysis.py 明阳电路
  python scripts/gap_analysis.py 明阳电路 --output 缺口报告.md
"""

import os
import re
import sys
import openpyxl
import requests
import fitz  # PyMuPDF

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from src.qdrant_ops import search_points, count_points

CLIENT_DATA_ROOT = r"C:\Users\mingh\client-data\raw\客户档案"
LLM_API_KEY = "sk-340ed7819c2346508c0a46a80df85999"
LLM_BASE_URL = "https://api.deepseek.com/v1"
LLM_MODEL = "deepseek-chat"

# 已知模块关键词
MODULE_PATTERNS = [
    "供应商管理", "寻源管理", "采购订单", "合同管理", "价格管理",
    "招投标", "竞价", "询价", "预算管理", "采购申请", "商城",
    "结算", "付款", "质量", "风控", "供应商协同", "会员",
    "主数据", "价格库", "工作流", "业务规则", "消息管理",
    "移动端", "门户", "报表", "BI", "审批",
]


def find_client_dir(name: str) -> str | None:
    if not os.path.isdir(CLIENT_DATA_ROOT):
        return None
    for n in os.listdir(CLIENT_DATA_ROOT):
        if name in n and os.path.isdir(os.path.join(CLIENT_DATA_ROOT, n)):
            return os.path.join(CLIENT_DATA_ROOT, n)
    return None


# ─── 数据格式解析 ───────────────────────────────────────────

def extract_from_xlsx_master(filepath: str) -> list:
    """从客户主数据xlsx提取'购买模块'列（最高优先级）"""
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        # 找"购买模块"列
        for col_i, h in enumerate(headers):
            if h and "购买模块" in str(h):
                for row in ws.iter_rows(min_row=2, values_only=True):
                    val = row[col_i] if col_i < len(row) else None
                    if val:
                        text = str(val).strip()
                        # 逗号/顿号/分号分隔
                        parts = re.split(r'[,，；;、\n]', text)
                        for p in parts:
                            p = p.strip()
                            if p and len(p) > 1:
                                yield p
                break
    except Exception as e:
        print(f"    xlsx读取失败: {e}")


def extract_from_doc(filepath: str) -> str:
    """用win32com读取Word文档，支持.doc和.docx，30秒超时"""
    import win32com.client
    import pythoncom
    import threading

    text_parts = []
    result = {"done": False, "error": None}

    def _read():
        try:
            pythoncom.CoInitialize()  # 初始化COM
            try:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                word.DisplayAlerts = False
                try:
                    doc = word.Documents.Open(os.path.abspath(filepath), ReadOnly=True)
                    try:
                        count = 0
                        for para in doc.Paragraphs:
                            if count >= 500:
                                break
                            t = para.Range.Text.strip()
                            if t:
                                text_parts.append(t)
                                count += 1
                        for i, table in enumerate(doc.Tables):
                            if i >= 3:
                                break
                            for row in table.Rows:
                                for cell in row.Cells:
                                    t = cell.Range.Text.strip()
                                    if t:
                                        text_parts.append(t)
                    finally:
                        doc.Close(False)
                finally:
                    word.Quit()
            finally:
                pythoncom.CoUninitialize()
            result["done"] = True
        except Exception as e:
            result["error"] = str(e)
            result["done"] = True

    t = threading.Thread(target=_read, daemon=True)
    t.start()
    t.join(timeout=30)  # 30秒超时
    if not result["done"]:
        print(f"    DOC读取超时（>30s），跳过: {os.path.basename(filepath)}")
        return ""
    if result["error"]:
        print(f"    Word读取失败: {result['error']}")
    return "\n".join(text_parts)


def extract_from_pdf(filepath: str) -> str:
    """用PyMuPDF提取PDF文字"""
    text_parts = []
    try:
        doc = fitz.open(filepath)
        for page in doc:
            t = page.get_text()
            if t.strip():
                text_parts.append(t.strip())
        doc.close()
    except Exception as e:
        print(f"    PDF读取失败: {e}")
    return "\n".join(text_parts)


def extract_modules_from_text(text: str) -> list:
    """从文本内容提取模块关键词"""
    found = []
    for pat in MODULE_PATTERNS:
        if pat in text:
            found.append(pat)
    # 去重
    seen = set()
    result = []
    for f in found:
        if f not in seen:
            seen.add(f)
            result.append(f)
    return result


# ─── 多数据源整合 ────────────────────────────────────────────

def _get_last_year() -> int:
    from datetime import datetime
    return datetime.now().year - 1


def _read_workorder_batch(filepath: str, year: int) -> list:
    import openpyxl as opx
    records = []
    try:
        wb = opx.load_workbook(filepath, data_only=True)
        for ws in wb.worksheets:
            hdr_row = next(ws.iter_rows(min_row=1, max_row=1))
            headers = [c.value for c in hdr_row]
            col_idx = {str(h).strip(): i for i, h in enumerate(headers) if h}
            required = ["标题", "描述", "解决方案", "根本原因", "模块", "提单时间"]
            if not all(k in col_idx for k in required):
                continue
            for row in ws.iter_rows(min_row=2, values_only=True):
                rd = {h: (row[i] if i < len(row) else None) for h, i in col_idx.items()}
                t = rd.get("提单时间", "")
                if not t:
                    continue
                record_year = None
                if hasattr(t, 'year'):
                    record_year = t.year
                else:
                    from datetime import datetime
                    ts = str(t).strip()[:10]
                    for fmt in ["%Y-%m-%d", "%Y/%m/%d"]:
                        try:
                            record_year = datetime.strptime(ts, fmt).year
                            break
                        except:
                            continue
                if record_year != year:
                    continue
                desc = str(rd.get("描述") or "").strip()
                solution = str(rd.get("解决方案") or "").strip()
                if not desc and not solution:
                    continue
                records.append({
                    "标题": str(rd.get("标题") or "").strip(),
                    "描述": desc,
                    "解决方案": solution,
                    "根本原因": str(rd.get("根本原因") or "").strip(),
                    "模块": str(rd.get("模块") or "").strip(),
                })
    except Exception as e:
        print(f"    读取工单失败: {e}")
    return records


def _llm_extract_from_workorders(records: list, batch_size: int = 10) -> dict:
    """基于产品功能清单，提取运维工单中的模块使用情况和功能使用情况。"""
    import requests, re, json

    # 加载产品功能层级
    hierarchy_path = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        "references", "product_modules_hierarchy.json"
    )
    with open(hierarchy_path, encoding="utf-8") as f:
        module_list = json.load(f)

    # 构建 prompt 中的模块清单（带序号）
    mod_lines = []
    for m in module_list:
        mod_lines.append(f"- {m['module']}：{', '.join(m['features'])}")
    module_catalog = "\n".join(mod_lines)

    all_used_modules = set()
    all_unused_modules = set()
    all_used_features = {}   # {模块名: [功能列表]}
    all_unused_features = {}  # {模块名: [功能列表]}

    for i in range(0, len(records), batch_size):
        batch = records[i:i + batch_size]
        items = []
        for j, rec in enumerate(batch, 1):
            items.append(
                "【工单{}】\n标题：{}\n描述：{}\n解决方案：{}\n根本原因：{}".format(
                    j,
                    rec.get("标题", ""),
                    rec.get("描述", "")[:300] if rec.get("描述") else "",
                    rec.get("解决方案", "")[:200] if rec.get("解决方案") else "",
                    rec.get("根本原因", "")[:150] if rec.get("根本原因") else ""
                )
            )
        block = "\n\n".join(items)

        prompt = (
            "你是SRM产品分析专家。基于以下运维工单，结合产品功能清单做分析。\n\n"
            "【产品功能清单】\n"
            + module_catalog +
            "\n\n"
            "【运维工单】\n"
            + block +
            "\n\n"
            "输出要求：\n"
            "1. 直接输出正文，不要任何开篇客套语\n"
            "2. 结构如下：\n\n"
            "## 模块使用情况\n"
            "【已用】：模块A、模块B...\n"
            "【未用】：模块C、模块D...\n\n"
            "## 已用模块功能分析\n"
            "### 模块A\n"
            "  【已用功能】：功能a、功能b\n"
            "  【未用功能】：功能c、功能d\n"
            "### 模块B\n"
            "  【已用功能】：功能x\n"
            "  【未用功能】：功能y、功能z\n"
            "...\n\n"
            "请开始分析："
        )

        try:
            resp = requests.post(
                LLM_BASE_URL + "/chat/completions",
                headers={"Authorization": "Bearer " + LLM_API_KEY, "Content-Type": "application/json"},
                json={"model": LLM_MODEL, "messages": [{"role": "user", "content": prompt}],
                      "temperature": 0.1, "max_tokens": 3000},
                timeout=180
            )
            if resp.status_code != 200:
                print(f"    LLM API错误 {resp.status_code}: {resp.text[:200]}")
                continue

            text = resp.json()["choices"][0]["message"]["content"]

            # 解析 "【已用】"
            m_used = re.search(r"【已用】\s*[:：]?\s*(.+?)(?:\n|【未用】|$)", text, re.DOTALL)
            if m_used:
                for mod in re.findall(r'[^，,\s、；;]+', m_used.group(1)):
                    all_used_modules.add(mod.strip())

            # 解析 "【未用】"
            m_unused = re.search(r"【未用】\s*[:：]?\s*(.+?)(?:\n##|\Z)", text, re.DOTALL)
            if m_unused:
                for mod in re.findall(r'[^，,\s、；;]+', m_unused.group(1)):
                    all_unused_modules.add(mod.strip())

            # 解析 "## 已用模块功能分析" 块
            func_block = re.search(r"## 已用模块功能分析\s*\n(.+)", text, re.DOTALL)
            if func_block:
                func_text = func_block.group(1)
                # 匹配每个 ### 模块 块
                for mod_match in re.finditer(r"###\s*(.+?)\s*\n(.+?)(?=###|\Z)", func_text, re.DOTALL):
                    mod_name = mod_match.group(1).strip()
                    content = mod_match.group(2)

                    used_f = re.search(r"【已用功能】\s*[:：]?\s*(.+?)(?:\n|【未用功能】|$)", content, re.DOTALL)
                    unused_f = re.search(r"【未用功能】\s*[:：]?\s*(.+?)(?:\n|$)", content, re.DOTALL)

                    if used_f:
                        feats = [f.strip() for f in re.findall(r'[^，,\s、；;]+', used_f.group(1)) if f.strip()]
                        if feats:
                            all_used_features[mod_name] = all_used_features.get(mod_name, []) + feats
                    if unused_f:
                        feats = [f.strip() for f in re.findall(r'[^，,\s、；;]+', unused_f.group(1)) if f.strip()]
                        if feats:
                            all_unused_features[mod_name] = all_unused_features.get(mod_name, []) + feats

        except Exception as e:
            print(f"    LLM调用异常: {e}")

        print(f"    已处理 {min(i + batch_size, len(records))}/{len(records)} 条工单")

    # 去重
    for k in all_used_features:
        all_used_features[k] = list(dict.fromkeys(all_used_features[k]))
    for k in all_unused_features:
        all_unused_features[k] = list(dict.fromkeys(all_unused_features[k]))

    return {
        "已用模块": sorted(all_used_modules),
        "未用模块": sorted(all_unused_modules),
        "已用功能": dict(all_used_features),
        "未用功能": dict(all_unused_features),
    }


def _extract_workorder_summary(client_dir: str, year: int) -> dict:
    """返回供报告使用的工单分析结果字典（仅包含list值，兼容modules_by_source）。"""
    import os
    ops_dir = os.path.join(client_dir, "运维工单")
    if not os.path.isdir(ops_dir):
        return {}
    all_records = []
    xlsx_files = sorted([f for f in os.listdir(ops_dir) if f.endswith(".xlsx")])
    print(f"  扫描运维工单: {len(xlsx_files)} 个文件，筛选{year}年数据")
    for fn in xlsx_files:
        recs = _read_workorder_batch(os.path.join(ops_dir, fn), year)
        if recs:
            print(f"    {fn}: {len(recs)} 条有效记录")
            all_records.extend(recs)
    if not all_records:
        print(f"  无{year}年工单数据")
        return {}
    print(f"  共{len(all_records)}条工单，LLM模块分析中...")
    summary = _llm_extract_from_workorders(all_records)

    used = summary["已用模块"]
    unused = summary["未用模块"]
    used_str = "、".join(used) if used else "无"
    unused_str = "、".join(unused) if unused else "无"

    # 保存到全局变量，供 main() 中的报告生成使用
    global _last_workorder_summary
    _last_workorder_summary = summary

    return {
        f"运维工单({year}年)": [
            f"共{len(all_records)}条工单",
            f"已用模块：{used_str}",
            f"未用模块：{unused_str}",
        ],
    }


# 全局变量，存储最近一次工单分析的完整结果（供 main() 使用）
_last_workorder_summary = None


def _extract_from_workorder_simple(filepath: str) -> list:
    import openpyxl as opx
    try:
        wb = opx.load_workbook(filepath, data_only=True)
        for ws in wb.worksheets:
            headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            for col_i, h in enumerate(headers):
                if h and any(k in str(h) for k in ["模块", "产品", "系统", "分类"]):
                    mods = []
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        val = row[col_i] if col_i < len(row) else None
                        if val and len(str(val).strip()) > 1:
                            mods.append(str(val).strip())
                    if mods:
                        return mods
    except:
        pass
    return []


def extract_all_modules(client_dir: str) -> dict:
    import os
    result = {}
    year = _get_last_year()
    # 1. 客户主数据xlsx
    for subdir in ["基础数据", "客户档案", "其他文档"]:
        sub = os.path.join(client_dir, subdir)
        if not os.path.isdir(sub):
            continue
        for fn in os.listdir(sub):
            if "客户主数据" in fn and fn.endswith(".xlsx"):
                print(f"  读取主数据: {fn}")
                mods = list(extract_from_xlsx_master(os.path.join(sub, fn)))
                if mods:
                    result["已购模块(xlsx)"] = mods
    # 2. 运维工单（LLM语义分析）
    ops_summary = _extract_workorder_summary(client_dir, year)
    if ops_summary:
        result.update(ops_summary)
    # 3. 蓝图方案目录
    blueprint_dir = os.path.join(client_dir, "蓝图方案")
    if os.path.isdir(blueprint_dir):
        pdf_modules, doc_modules = [], []
        for fn in sorted(os.listdir(blueprint_dir)):
            fp = os.path.join(blueprint_dir, fn)
            if fn.endswith(".pdf"):
                print(f"  读取PDF: {fn[:40]}")
                text = extract_from_pdf(fp)
                if text:
                    pdf_modules.extend(extract_modules_from_text(text))
            elif fn.endswith(".doc") or fn.endswith(".docx"):
                print(f"  读取DOC: {fn[:40]}")
                text = extract_from_doc(fp)
                if text:
                    doc_modules.extend(extract_modules_from_text(text))
        if pdf_modules:
            result["蓝图PDF"] = _dedup_list(pdf_modules)
        if doc_modules:
            result["蓝图DOC"] = _dedup_list(doc_modules)
    # 4. 商务信息目录
    biz_dir = os.path.join(client_dir, "商务信息")
    if os.path.isdir(biz_dir):
        biz_modules = []
        for fn in os.listdir(biz_dir):
            if fn.endswith(".xlsx"):
                biz_modules.extend(_extract_from_workorder_simple(os.path.join(biz_dir, fn)))
        if biz_modules:
            result["合同附件"] = _dedup_list(biz_modules)
    return result


def _dedup_list(lst: list) -> list:
    seen = set()
    result = []
    for x in lst:
        if x not in seen:
            seen.add(x)
            result.append(x)
    return result


def extract_all_modules(client_dir: str) -> dict:
    """返回 {来源: [内容列表]}"""
    import os
    result = {}
    year = _get_last_year()

    # 1. 客户主数据xlsx（最高优先级）
    for subdir in ["基础数据", "客户档案", "其他文档"]:
        sub = os.path.join(client_dir, subdir)
        if not os.path.isdir(sub):
            continue
        for fn in os.listdir(sub):
            if "客户主数据" in fn and fn.endswith(".xlsx"):
                print(f"  读取主数据: {fn}")
                mods = list(extract_from_xlsx_master(os.path.join(sub, fn)))
                if mods:
                    result["已购模块(xlsx)"] = mods

    # 2. 运维工单（LLM语义分析）
    ops_summary = _extract_workorder_summary(client_dir, year)
    if ops_summary:
        result.update(ops_summary)

    # 3. 蓝图方案目录
    blueprint_dir = os.path.join(client_dir, "蓝图方案")
    if os.path.isdir(blueprint_dir):
        pdf_modules = []
        doc_modules = []
        for fn in sorted(os.listdir(blueprint_dir)):
            fp = os.path.join(blueprint_dir, fn)
            if fn.endswith(".pdf"):
                print(f"  读取PDF: {fn[:40]}")
                text = extract_from_pdf(fp)
                if text:
                    mods = extract_modules_from_text(text)
                    pdf_modules.extend(mods)
            elif fn.endswith(".doc") or fn.endswith(".docx"):
                print(f"  读取DOC: {fn[:40]}")
                text = extract_from_doc(fp)
                if text:
                    mods = extract_modules_from_text(text)
                    doc_modules.extend(mods)
        if pdf_modules:
            result["蓝图PDF"] = _dedup_list(pdf_modules)
        if doc_modules:
            result["蓝图DOC"] = _dedup_list(doc_modules)

    # 4. 商务信息目录
    biz_dir = os.path.join(client_dir, "商务信息")
    if os.path.isdir(biz_dir):
        biz_modules = []
        for fn in os.listdir(biz_dir):
            if fn.endswith(".xlsx"):
                mods = _extract_from_workorder_simple(os.path.join(biz_dir, fn))
                biz_modules.extend(mods)
        if biz_modules:
            result["合同附件"] = _dedup_list(biz_modules)

    return result


def _dedup_list(lst: list) -> list:
    seen = set()
    result = []
    for x in lst:
        if x not in seen:
            seen.add(x)
            result.append(x)
    return result


# ─── LLM 报告生成 ────────────────────────────────────────────

def call_llm(messages: list) -> str:
    resp = requests.post(
        f"{LLM_BASE_URL}/chat/completions",
        headers={"Authorization": f"Bearer {LLM_API_KEY}", "Content-Type": "application/json"},
        json={"model": LLM_MODEL, "messages": messages, "temperature": 0.3},
        timeout=180
    )
    if resp.status_code == 200:
        return resp.json()["choices"][0]["message"]["content"]
    return f"LLM failed: {resp.status_code}"


def generate_report(client_name: str, modules_by_source: dict,
                    product_results: list, output_path: str = None,
                    output_format: str = "md",
                    workorder_summary: dict = None) -> str:
    # 合并所有模块（排除dict类型的元数据）
    all_mods_by_src = []
    for src, mods in modules_by_source.items():
        if not isinstance(mods, list):
            continue
        for m in mods:
            all_mods_by_src.append(f"- {m}（来源：{src}）")
    used_text = "\n".join(all_mods_by_src) if all_mods_by_src else "（未能提取到模块信息）"

    # 如果有工单分析结果，补充已用/未用模块信息
    workorder_text = ""
    if workorder_summary:
        used = workorder_summary.get("已用模块", [])
        unused = workorder_summary.get("未用模块", [])
        if used:
            workorder_text += f"\n【运维工单分析 - 已用模块】：{'、'.join(used)}"
        if unused:
            workorder_text += f"\n【运维工单分析 - 未用模块】：{'、'.join(unused)}"

    prod_text = "\n".join([
        f"- {p[0].get('module','')}/{p[0].get('type','')}: {p[0].get('text','')[:80]}..."
        for p in product_results[:20]
    ])

    prompt = f"""你是SRM产品专家。为客户做功能缺口分析。

客户名称：{client_name}

已用/已购模块（多数据源）：
{used_text}
{workorder_text}

产品功能库（部分）：
{prod_text}

输出要求：
1. 直接输出正文，不要任何开篇客套语
2. 标题用##，加粗用**文字**
3. 输出结构：

## 1. 已用功能概览

## 2. 功能缺口
### 2.1 已购未充分使用
### 2.2 可挖掘机会

## 3. 优先级推荐（3-5个）

## 4. 总结"""

    report = call_llm([{"role": "user", "content": prompt}])
    # 去掉开篇客套语（如LLM说"好的，作为..."）
    lines = report.split('\n')
    skip_patterns = ['好的，', '作为SRM产品专家', '以下是', '我将基于', '下面为', '以下为']
    filtered = []
    for line in lines:
        stripped = line.strip()
        if stripped and not stripped.startswith('#') and any(stripped.startswith(p) for p in skip_patterns):
            continue
        filtered.append(line)
    report = '\n'.join(filtered)

    if output_path:
        if output_format == "docx":
            # 先写md临时文件，再转docx
            md_path = output_path.replace(".docx", ".md")
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(f"# {client_name} - 功能缺口分析\n\n**时间**: 2026-03-19\n\n{report}")
            # 调用md2docx转换
            try:
                from scripts.md2docx import convert_markdown_to_docx
                ok = convert_markdown_to_docx(md_path, output_path)
                if ok:
                    print(f"Report saved: {output_path}")
                else:
                    print(f"Report saved (md): {md_path}")
            except Exception as e:
                print(f"DOCX conversion failed: {e}, saved as md: {md_path}")
        else:
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(f"# {client_name} - 功能缺口分析\n\n**时间**: 2026-03-19\n\n{report}")
            print(f"Report saved: {output_path}")
    return report


# ─── 主入口 ────────────────────────────────────────────────

def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("客户名称")
    parser.add_argument("--output", "-o", help="Output file path")
    parser.add_argument("--format", "-f", choices=["md", "docx"], default="md",
                        help="Output format: md or docx (default: md)")
    args = parser.parse_args()

    client_name = args.客户名称
    print(f"\n[search] Client: {client_name}")

    client_dir = find_client_dir(client_name)
    if not client_dir:
        print(f"[ERROR] Client dir not found: {CLIENT_DATA_ROOT}")
        sys.exit(1)
    print(f"[OK] Dir: {client_dir}")

    print("\n[import] Extracting data sources...")
    modules_by_source = extract_all_modules(client_dir)
    for src, mods in modules_by_source.items():
        print(f"   {src}: {len(mods)} modules")

    if not modules_by_source:
        print("[WARN] No modules extracted")
        sys.exit(1)

    total = count_points()
    if total == 0:
        print("[ERROR] Knowledge base empty. Run: python scripts/import_knowledge.py")
        sys.exit(1)

    print(f"\n[search] Searching product DB ({total} records)...")
    # 用已提取的模块名构建检索query
    all_mod_names = sum([m for m in modules_by_source.values()], [])
    query = " ".join(all_mod_names[:15]) if len(" ".join(all_mod_names[:15])) > 10 else "SRM采购管理供应商合同订单"
    results = search_points(query, top_k=50)
    print(f"   Found {len(results)} product features")

    # 提取所有模块名列表（排除dict类型的元数据）
    all_mod_names = sum([m for m in modules_by_source.values() if isinstance(m, list)], [])
    query = " ".join(all_mod_names[:15]) if len(" ".join(all_mod_names[:15])) > 10 else "SRM采购管理供应商合同订单"

    print(f"  已用模块: {len(_last_workorder_summary['已用模块']) if _last_workorder_summary else 0}")
    print(f"  未用模块: {len(_last_workorder_summary['未用模块']) if _last_workorder_summary else 0}")

    print(f"\n[AI] Generating analysis report...")
    report = generate_report(client_name, modules_by_source, results,
                            args.output, args.format,
                            workorder_summary=_last_workorder_summary)

    print(f"\n{'='*60}\nGap Analysis Report\n{'='*60}\n{report}")


if __name__ == "__main__":
    main()
