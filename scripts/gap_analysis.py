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



def _llm_extract_from_workorders(records: list, batch_size: int = None, bought_modules: list = None) -> dict:
    """三相方案：Phase1模块分类 + Phase2A子功能 + Phase2B障碍分析。

    bought_modules: 合同买了的模块列表（来自主数据），可None。
    返回：{已用模块, 未用模块, 买了没用模块, 已用功能, 未用功能, 障碍分析}
    """
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "_llm_workorder_phases",
        os.path.join(os.path.dirname(__file__), "_llm_workorder_phases.py")
    )
    phases_mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(phases_mod)

    # 直接调用三个 phase
    hierarchy = phases_mod.load_hierarchy()

    # 复用 build_block
    def build_block(recs):
        items = []
        for j, rec in enumerate(recs, 1):
            items.append(
                "[工单%d]\n标题：%s\n描述：%s" % (
                    j,
                    rec.get("标题", ""),
                    (rec.get("描述") or "")[:200]
                )
            )
        return "\n\n".join(items)

    block = build_block(records)

    # 标准模块名列表（用于 Phase1 prompt）
    std_names = [item["module"] for item in hierarchy]
    mod_catalog = "\n".join(["- %s" % n for n in std_names])

    # ── Phase 1 ──────────────────────────────────────────────
    prompt_p1 = (
        "你是SRM产品分析专家。基于运维工单，判断客户在每个模块的使用情况。\n\n"
        "【产品功能模块清单】\n"
        + mod_catalog + "\n\n"
        "【运维工单】\n"
        + block + "\n\n"
        "输出要求：\n"
        "1. 直接输出正文，不要任何开篇客套语\n"
        "2. 只判断模块，不分析子功能\n"
        "3. 判断标准：\n"
        "   - 【已用】：买了这个模块，且工单中频繁出现，说明实际在用\n"
        "   - 【买了没用】：买了这个模块，但工单中没有或极少出现，说明买了没真正用起来\n"
        "   - 【未用】：没买这个模块，或者买了但工单中完全没有涉及\n"
        "4. 结构如下：\n\n"
        "## 模块使用情况\n"
        "【已用】：<模块列表>\n"
        "【买了没用】：<模块列表>\n"
        "【未用】：<模块列表>\n\n"
        "请开始分析"
    )

    print("    [Phase1] 判断三分类...")
    raw_p1 = call_llm([{"role": "user", "content": prompt_p1}])

    def extract_cat(text, key):
        for prefix in ["【%s】：" % key, "【%s】: " % key]:
            idx = text.rfind(prefix)
            if idx == -1:
                continue
            seg = text[idx + len(prefix):]
            end = len(seg)
            for p in ["【", "##"]:
                e = seg.find(p)
                if e != -1 and e < end:
                    end = e
            raw = seg[:end].strip()
            import re as _re
            items = _re.split(r'[、，,;；\n]+', raw)
            result = [it.strip().strip('，。、.').strip() for it in items if it.strip() and len(it.strip()) > 1]
            if result:
                return result
        return []

    used = extract_cat(raw_p1, "已用")
    bought_unused = extract_cat(raw_p1, "买了没用")
    unused = extract_cat(raw_p1, "未用")
    print("    Phase1完成: 已用(%d) 买了没用(%d) 未用(%d)" % (len(used), len(bought_unused), len(unused)))

    # ── Phase 2A ─────────────────────────────────────────────
    used_funcs = {}
    if used:
        target = {item["module"]: item.get("features", []) for item in hierarchy if item["module"] in used}
        func_cat = "\n".join(["- %s：%s" % (n, ", ".join(fs)) for n, fs in target.items()])
        prompt_p2a = (
            "你是SRM产品分析专家。基于运维工单，分析已用模块的子功能使用情况。\n\n"
            "【已用模块及其子功能清单】\n"
            + func_cat + "\n\n"
            "【运维工单】\n"
            + block + "\n\n"
            "输出要求：\n"
            "1. 直接输出正文，不要任何开篇客套语\n"
            "2. 针对每个已用模块，列出工单中实际出现的具体功能（只列工单里有的，不要编造）\n"
            "3. 结构如下：\n\n"
            "## 已用模块功能分析\n"
            "### <模块名>\n"
            "  【已用功能】：<工单中实际用到的具体功能>\n"
            "  【未用功能】：<该模块下工单中完全没有涉及的功能>\n"
            "...（每个已用模块都要分析）\n\n"
            "请开始分析"
        )
        print("    [Phase2A] 深挖%d个已用模块子功能..." % len(used))
        raw_p2a = call_llm([{"role": "user", "content": prompt_p2a}])
        import re as _re2
        cur_mod = None
        for line in raw_p2a.split('\n'):
            line_s = line.strip()
            if not line_s:
                continue
            m = _re2.match(r'#{1,3}\s*(.+?)\s*$', line)
            if m:
                cur_mod = m.group(1).strip()
                if cur_mod not in used_funcs:
                    used_funcs[cur_mod] = {"已用": [], "未用": []}
                continue
            if not cur_mod:
                continue
            for p, key in [("【已用功能】：", "已用"), ("【未用功能】：", "未用")]:
                if p in line:
                    items = _re2.split(r'[、，,;；\n]+', line.split(p)[-1])
                    for f in items:
                        f = f.strip().strip('，。、.').strip()
                        if f and len(f) > 1:
                            used_funcs[cur_mod][key].append(f)

    # ── Phase 2B ─────────────────────────────────────────────
    barrier_results = {}
    if bought_unused:
        target = {item["module"]: item.get("features", []) for item in hierarchy if item["module"] in bought_unused}
        barrier_cat = "\n".join(["- %s：%s" % (n, ", ".join(fs)) for n, fs in target.items()])
        prompt_p2b = (
            "你是SRM产品分析专家。基于运维工单，分析以下模块为什么买了但没有实际使用。\n\n"
            "【买了没用的模块及其子功能】\n"
            + barrier_cat + "\n\n"
            "【运维工单】\n"
            + block + "\n\n"
            "输出要求：\n"
            "1. 直接输出正文，不要任何开篇客套语\n"
            "2. 分析每个模块没被使用的原因/障碍\n"
            "3. 结构如下：\n\n"
            "## 买了没用模块分析\n"
            "### <模块名>\n"
            "  【障碍】：<为什么没用起来>\n"
            "  【可挖掘机会】：<如果启用，能解决什么问题>\n"
            "...（每个买了没用的模块都要分析）\n\n"
            "请开始分析"
        )
        print("    [Phase2B] 分析%d个买了没用模块障碍..." % len(bought_unused))
        raw_p2b = call_llm([{"role": "user", "content": prompt_p2b}])
        import re as _re3
        cur_mod = None
        for line in raw_p2b.split('\n'):
            line_s = line.strip()
            if not line_s:
                continue
            m = _re3.match(r'#{1,3}\s*(.+?)\s*$', line)
            if m:
                cur_mod = m.group(1).strip()
                if cur_mod not in barrier_results:
                    barrier_results[cur_mod] = {"障碍": [], "机会": []}
                continue
            if not cur_mod:
                continue
            for p, key in [("【障碍】：", "障碍"), ("【原因】：", "障碍"),
                            ("【可挖掘机会】：", "机会"), ("【机会】：", "机会")]:
                if p in line:
                    items = _re3.split(r'[、，,;；\n]+', line.split(p)[-1])
                    for f in items:
                        f = f.strip().strip('，。、.').strip()
                        if f and len(f) > 1:
                            barrier_results[cur_mod][key].append(f)

    return {
        "已用模块": sorted(set(used)),
        "未用模块": sorted(set(unused)),
        "买了没用模块": sorted(set(bought_unused)),
        "已用功能": used_funcs,
        "未用功能": {},
        "障碍分析": barrier_results,
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

    # === Prompt v2：正文不加粗，只用##结构 ===
    # 回滚方案：注释下面两行，改回 v1 prompt 即可
    prompt_v2 = f"""你是SRM产品专家。为客户做功能缺口分析。

客户名称：{client_name}

已用/已购模块（多数据源）：
{used_text}
{workorder_text}

产品功能库（部分）：
{prod_text}

输出要求：
1. 直接输出正文，不要任何开篇客套语
2. 标题用##，描述内容不加粗（不用**文字**），不强调，不突出显示
3. 输出结构：

## 1. 已用功能概览

## 2. 功能缺口
### 2.1 已购未充分使用
### 2.2 可挖掘机会

## 3. 优先级推荐（3-5个）

## 4. 总结"""

    # === Prompt v1（回滚用）：允许加粗 ===
    # prompt_v1 = f"""你是SRM产品专家。为客户做功能缺口分析...
    # 输出要求：2. 标题用##，加粗用**文字**
    # """

    report = call_llm([{"role": "user", "content": prompt_v2}])

    # 清洗：去掉开篇客套语
    lines = report.split('\n')
    skip_patterns = ['好的，', '作为SRM产品专家', '以下是', '我将基于', '下面为', '以下为']
    filtered = []
    for line in lines:
        stripped = line.strip()
        if stripped and not stripped.startswith('#') and any(stripped.startswith(p) for p in skip_patterns):
            continue
        filtered.append(line)
    # 清洗（回退保险）：全文统一去掉所有加粗
    import re
    report = re.sub(r'\*\*([^*]+)\*\*', r'\1', report)

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
