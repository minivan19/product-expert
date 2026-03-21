#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
场景1：需求 → 功能推荐
用法：
  python scripts/search_features.py "客户希望管理供应商资质有效期"
  python scripts/search_features.py "客户希望管理供应商资质有效期" --output /path/to/report.md
"""

import os
import sys
import argparse
from datetime import datetime
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from src.qdrant_ops import search_points, count_points

LLM_API_KEY = "sk-340ed7819c2346508c0a46a80df85999"
LLM_BASE_URL = "https://api.deepseek.com/v1"
LLM_MODEL = "deepseek-chat"
SEARCH_TOP_K = 20
OUTPUT_ROOT = "/Users/limingheng/AI/client-data/产品标准推荐"


def call_llm(messages: list) -> str:
    import requests
    url = f"{LLM_BASE_URL}/chat/completions"
    resp = requests.post(
        url,
        headers={"Authorization": f"Bearer {LLM_API_KEY}", "Content-Type": "application/json"},
        json={"model": LLM_MODEL, "messages": messages, "temperature": 0.3},
        timeout=120
    )
    if resp.status_code == 200:
        return resp.json()["choices"][0]["message"]["content"]
    return f"LLM调用失败: {resp.status_code}"


def format_results(results: list) -> str:
    lines = []
    for i, (payload, score) in enumerate(results, 1):
        lines.append(
            f"[{i}] 相关度:{score:.3f}\n"
            f"模块: {payload.get('module','')} | "
            f"类型: {payload.get('type','')}（来源:{payload.get('source','')}）\n"
            f"文档: {payload.get('doc_name','')}\n"
            f"内容: {payload.get('text','')[:200]}\n---"
        )
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="场景1：需求 → 功能推荐")
    parser.add_argument("need", help="客户需求描述")
    parser.add_argument("--output", "-o", help="输出Markdown文件路径（默认保存到产品标准推荐目录）")
    args = parser.parse_args()

    need = args.need
    total = count_points()
    if total == 0:
        print("Warning: 知识库为空，请先运行: python scripts/import_knowledge.py")
        sys.exit(1)

    print(f"\nQuery: {need}")
    print(f"Knowledge base: {total} records")
    results = search_points(need, top_k=SEARCH_TOP_K)
    print(f"Found {len(results)} results")

    if not results:
        print("⚠️  未找到匹配功能")
        sys.exit(0)

    print("LLM analyzing...")
    system_prompt = """你是SRM产品专家。基于检索到的产品功能，为用户需求推荐功能。

输出格式：
1. 先理解用户需求
2. 推荐5-10个最相关的功能
3. 每条：功能名称 + 匹配理由 + 来源文档 + 适用场景
4. 按相关度排序"""

    user_prompt = f"""用户需求：{need}

检索结果：
{format_results(results)}

请推荐功能。"""

    rec = call_llm([{"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}])

    output_content = f"# 功能推荐报告\n\n**需求：** {need}\n\n{rec}"

    if args.output:
        output_path = args.output
    else:
        os.makedirs(OUTPUT_ROOT, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(OUTPUT_ROOT, f"需求推荐_{ts}.md")

    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(output_content)
        print(f"\n报告已保存: {output_path}")
    except Exception as e:
        print(f"\n报告保存失败: {e}")
        print(f"\n=== Feature Recommendations ===\n{rec}")


if __name__ == "__main__":
    main()
