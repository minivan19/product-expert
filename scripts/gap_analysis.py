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

def extract_all_modules(client_dir: str) -> dict:
    """
    返回 {来源: [模块列表]}
    按优先级处理各数据源
    """
    result = {}

    # 1. 客户主数据xlsx（最高优先级："购买模块"列）
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
                    print(f"    提取到: {mods}")

    # 2. 运维工单（已有逻辑）
    ops_dir = os.path.join(client_dir, "运维工单")
    if os.path.isdir(ops_dir):
        ops_modules = []
        for fn in os.listdir(ops_dir):
            if fn.endswith(".xlsx"):
                mods = _extract_from_workorder(os.path.join(ops_dir, fn))
                ops_modules.extend(mods)
        if ops_modules:
            result["运维工单"] = _dedup_list(ops_modules)

    # 3. 蓝图方案目录：PDF流程图 + DOC手册
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
                    if mods:
                        print(f"    提取到模块: {mods}")
            elif fn.endswith(".doc") or fn.endswith(".docx"):
                print(f"  读取DOC: {fn[:40]}")
                text = extract_from_doc(fp)
                if text:
                    mods = extract_modules_from_text(text)
                    doc_modules.extend(mods)
                    if mods:
                        print(f"    提取到模块: {mods}")
        if pdf_modules:
            result["蓝图PDF"] = _dedup_list(pdf_modules)
        if doc_modules:
            result["蓝图DOC"] = _dedup_list(doc_modules)

    # 4. 商务信息目录（合同附件xlsx）
    biz_dir = os.path.join(client_dir, "商务信息")
    if os.path.isdir(biz_dir):
        biz_modules = []
        for fn in os.listdir(biz_dir):
            if fn.endswith(".xlsx"):
                mods = _extract_from_workorder(os.path.join(biz_dir, fn))
                biz_modules.extend(mods)
        if biz_modules:
            result["合同附件"] = _dedup_list(biz_modules)

    return result


def _extract_from_workorder(filepath: str) -> list:
    """从运维/合同xlsx提取模块列"""
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
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
                    output_format: str = "md") -> str:
    # 合并所有模块
    all_mods_by_src = []
    for src, mods in modules_by_source.items():
        for m in mods:
            all_mods_by_src.append(f"- {m}（来源：{src}）")
    used_text = "\n".join(all_mods_by_src) if all_mods_by_src else "（未能提取到模块信息）"

    prod_text = "\n".join([
        f"- {p[0].get('module','')}/{p[0].get('type','')}: {p[0].get('text','')[:80]}..."
        for p in product_results[:20]
    ])

    prompt = f"""你是SRM产品专家。为客户做功能缺口分析。

客户名称：{client_name}

已用/已购模块（多数据源）：
{used_text}

产品功能库（部分）：
{prod_text}

请输出：
## 客户：{client_name}
## 数据来源：{', '.join(modules_by_source.keys()) if modules_by_source else '无'}

## 1. 已用功能概览（综合所有数据源）

## 2. 功能缺口
### 2.1 已购未充分使用
### 2.2 可挖掘机会

## 3. 优先级推荐（3-5个）

## 4. 总结"""

    report = call_llm([{"role": "user", "content": prompt}])
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

    print("\n[AI] Generating analysis report...")
    report = generate_report(client_name, modules_by_source, results,
                            args.output, args.format)

    print(f"\n{'='*60}\nGap Analysis Report\n{'='*60}\n{report}")


if __name__ == "__main__":
    main()
