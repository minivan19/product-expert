#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
场景1：需求 → 功能推荐
输入客户需求描述，输出功能推荐清单（含理由和来源）
"""

import os
import sys
import json
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from qdrant_client import search, count_points, COLLECTION_NAME

# LLM配置（使用gptsapi.net）
LLM_API_KEY = "sk-Pzt8a346e78b733bfead64b269317c033e97cd59abfWoqEt"
LLM_BASE_URL = "https://api.gptsapi.net/v1"
LLM_MODEL = "gpt-3.5-turbo"

SEARCH_TOP_K = 20  # 向量检索返回的候选数量


def call_llm(messages: list) -> str:
    """调用LLM生成分析结果"""
    url = f"{LLM_BASE_URL}/chat/completions"
    headers = {
        "Authorization": f"Bearer {LLM_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": LLM_MODEL,
        "messages": messages,
        "temperature": 0.3
    }
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=120)
        if response.status_code == 200:
            return response.json()["choices"][0]["message"]["content"]
        else:
            return f"LLM调用失败: {response.status_code} - {response.text}"
    except Exception as e:
        return f"LLM调用异常: {e}"


def format_results_for_llm(results: list) -> str:
    """将检索结果格式化为LLM输入"""
    lines = []
    for i, (payload, score) in enumerate(results, 1):
        module = payload.get("module", "")
        func_type = payload.get("type", "")
        source = payload.get("source", "")
        version = payload.get("version", "")
        doc = payload.get("doc_name", "")
        text = payload.get("text", "")[:300]  # 限制长度

        lines.append(f"""[{i}] 相关度:{score:.3f}
模块: {module}
类型: {func_type}（来源:{source}, 版本:{version}）
来源文档: {doc}
功能内容: {text}

---""")
    return "\n".join(lines)


def generate_recommendations(need: str, results: list) -> str:
    """调用LLM基于检索结果生成推荐清单"""
    context = format_results_for_llm(results)

    system_prompt = """你是一名SRM产品专家。根据检索到的产品功能信息，为用户的需求或痛点提供功能推荐清单。

输出要求：
1. 先理解用户需求
2. 从检索结果中匹配最相关的功能
3. 每条推荐说明：
   - **推荐功能**：[功能名称]
   - **匹配理由**：为什么这个功能能解决用户需求
   - **来源**：来自哪份文档（功能类型：标准功能/新增/修正）
   - **适用场景**：建议在什么阶段上线
4. 如检索结果无法满足需求，明确说明"暂无匹配功能，建议进一步描述需求"
5. 按相关度从高到低排序，输出5-10条推荐
6. 只输出推荐清单，不要额外废话"""

    user_prompt = f"""用户需求/痛点：
{need}

产品功能知识库检索结果：
{context}

请基于以上检索结果，为用户推荐合适的产品功能。"""

    return call_llm([
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ])


def main():
    if len(sys.argv) < 2:
        print("用法: python search_features.py <需求描述>")
        print("示例: python search_features.py \"客户希望管理供应商的资质有效期，到期前能自动提醒\"")
        sys.exit(1)

    need = sys.argv[1]

    # 检查知识库
    total = count_points()
    if total == 0:
        print("⚠️  知识库为空，请先运行: python import_knowledge.py")
        sys.exit(1)

    print(f"\n🔍 需求: {need}")
    print(f"📚 知识库检索中...（当前 {total} 条记录）")

    # 向量检索
    results = search(need, top_k=SEARCH_TOP_K)
    print(f"   找到 {len(results)} 条相关结果")

    if not results:
        print("⚠️  未找到匹配的功能，请尝试换一种描述方式")
        sys.exit(0)

    # LLM分析生成推荐
    print(f"\n🤖 LLM分析中...")
    recommendations = generate_recommendations(need, results)

    print(f"\n{'='*60}")
    print("功能推荐结果")
    print(f"{'='*60}")
    print(recommendations)


if __name__ == "__main__":
    main()
