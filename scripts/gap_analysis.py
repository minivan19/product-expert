#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
场景2：客户功能缺口分析
用法：
  python scripts/gap_analysis.py 明阳电路
  python scripts/gap_analysis.py 明阳电路 --output 缺口报告.md
"""

import os
import re
import sys
import openpyxl
import requests

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from src.qdrant_ops import search_points, count_points

CLIENT_DATA_ROOT = r"C:\Users\mingh\client-data\raw\客户档案"
LLM_API_KEY = "sk-Pzt8a346e78b733bfead64b269317c033e97cd59abfWoqEt"
LLM_BASE_URL = "https://api.gptsapi.net/v1"
LLM_MODEL = "gpt-3.5-turbo"


def find_client_dir(name: str) -> str | None:
    if not os.path.isdir(CLIENT_DATA_ROOT):
        return None
    for n in os.listdir(CLIENT_DATA_ROOT):
        if name in n and os.path.isdir(os.path.join(CLIENT_DATA_ROOT, n)):
            return os.path.join(CLIENT_DATA_ROOT, n)
    return None


def extract_from_xlsx(filepath: str) -> list:
    modules = []
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        for ws in wb.worksheets:
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            for col_i, h in enumerate(headers):
                if h and any(k in str(h) for k in ["模块", "产品", "功能", "系统"]):
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        val = row[col_i] if col_i < len(row) else None
                        if val and len(str(val).strip()) > 1:
                            modules.append(str(val).strip())
                    break
    except:
        pass
    return list(dict.fromkeys(modules))  # 去重保留顺序


def extract_from_text(filepath: str) -> list:
    modules = []
    try:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        patterns = [
            r"供应商管理|寻源管理|采购订单|合同管理|价格管理|招投标|竞价|询价",
            r"预算管理|采购申请|商城|结算|付款|质量|风控|供应商协同|会员",
            r"主数据|价格库|工作流|业务规则|消息管理|移动端|门户",
        ]
        for pat in patterns:
            for m in re.findall(pat, content):
                modules.append(m)
    except:
        pass
    return list(dict.fromkeys(modules))


def extract_from_workorders(client_dir: str) -> list:
    modules = []
    ops_dir = os.path.join(client_dir, "运维工单")
    if not os.path.isdir(ops_dir):
        return modules
    for fn in os.listdir(ops_dir):
        if not fn.endswith(".xlsx"):
            continue
        filepath = os.path.join(ops_dir, fn)
        found = extract_from_xlsx(filepath)
        modules.extend([f"{m}（工单）" for m in found])
    return list(dict.fromkeys(modules))


def extract_all_modules(client_dir: str) -> dict:
    """返回 {来源: [模块列表]}"""
    result = {}

    # 商务信息
    biz = os.path.join(client_dir, "商务信息")
    if os.path.isdir(biz):
        mods = []
        for fn in os.listdir(biz):
            fp = os.path.join(biz, fn)
            if fn.endswith(".xlsx"):
                mods.extend(extract_from_xlsx(fp))
            elif fn.endswith((".md", ".txt")):
                mods.extend(extract_from_text(fp))
        if mods:
            result["合同模块"] = mods[:20]

    # 运维工单
    ops = os.path.join(client_dir, "运维工单")
    if os.path.isdir(ops):
        mods = extract_from_workorders(client_dir)
        if mods:
            result["运维工单"] = mods[:20]

    # 蓝图/用户手册
    for fn in os.listdir(client_dir):
        if any(k in fn for k in ["蓝图", "实施", "用户手册"]):
            fp = os.path.join(client_dir, fn)
            if fn.endswith(".md"):
                mods = extract_from_text(fp)
                if mods:
                    result["蓝图/手册"] = mods[:20]

    return result


def call_llm(messages: list) -> str:
    resp = requests.post(
        f"{LLM_BASE_URL}/chat/completions",
        headers={"Authorization": f"Bearer {LLM_API_KEY}", "Content-Type": "application/json"},
        json={"model": LLM_MODEL, "messages": messages, "temperature": 0.3},
        timeout=120
    )
    if resp.status_code == 200:
        return resp.json()["choices"][0]["message"]["content"]
    return f"LLM失败: {resp.status_code}"


def generate_report(client_name: str, modules_by_source: dict, product_results: list, output_path: str = None) -> str:
    # 合并所有模块
    all_mods = []
    for src, mods in modules_by_source.items():
        for m in mods:
            all_mods.append(f"- {m}（来源：{src}）")

    used_text = "\n".join(all_mods) if all_mods else "（未能提取到模块信息）"

    prod_text = "\n".join([
        f"- {p[0].get('module','')}/{p[0].get('type','')}: {p[0].get('text','')[:80]}..."
        for p in product_results[:20]
    ])

    prompt = f"""你是SRM产品专家。为客户做功能缺口分析。

客户名称：{client_name}

已用模块（数据来源）：
{used_text}

产品功能库（部分）：
{prod_text}

请输出：
## 客户：{client_name}
## 数据来源：{', '.join(modules_by_source.keys()) if modules_by_source else '无'}

## 1. 已用功能概览

## 2. 功能缺口
### 2.1 已购未充分使用
### 2.2 可挖掘机会

## 3. 优先级推荐（3-5个）

## 4. 总结"""

    report = call_llm([{"role": "user", "content": prompt}])
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"# {client_name} - 功能缺口分析\n\n**时间**: 2026-03-19\n\n{report}")
        print(f"✅ 报告已保存: {output_path}")
    return report


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("客户名称")
    parser.add_argument("--output", "-o", help="输出文件")
    args = parser.parse_args()

    client_name = args.客户名称
    print(f"\n🔍 客户: {client_name}")

    client_dir = find_client_dir(client_name)
    if not client_dir:
        print(f"❌ 未找到客户目录: {CLIENT_DATA_ROOT}")
        sys.exit(1)
    print(f"✅ 目录: {client_dir}")

    # 提取模块
    print("\n📂 提取数据源...")
    modules_by_source = extract_all_modules(client_dir)
    for src, mods in modules_by_source.items():
        print(f"   {src}: {len(mods)} 个")

    if not modules_by_source:
        print("⚠️  未提取到任何模块")
        sys.exit(1)

    # 检索产品功能
    total = count_points()
    if total == 0:
        print("⚠️  知识库为空，请先运行: python scripts/import_knowledge.py")
        sys.exit(1)

    print(f"\n📚 检索产品功能库（{total} 条）...")
    query = " ".join(sum([m for m in modules_by_source.values()], [])[:10])
    if len(query) < 10:
        query = "SRM采购管理供应商合同订单"
    results = search_points(query, top_k=50)
    print(f"   找到 {len(results)} 条")

    # 生成报告
    print("\n🤖 生成分析报告...")
    report = generate_report(client_name, modules_by_source, results, args.output)

    print(f"\n{'='*60}\n缺口分析结果\n{'='*60}\n{report}")


if __name__ == "__main__":
    main()
