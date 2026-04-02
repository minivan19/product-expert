#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OpenViking operations for product-expert
操作 viking://resources/product（阿里 text-embedding-v4，2048维）
与 Mem0 的 openclaw_memories 完全隔离
"""

import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Optional

# OpenViking 资源 URI（产品知识库已通过 ov add-resource 导入）
PRODUCT_RESOURCE_URI = "viking://resources/srm-products"

# 向量维度（本地 Ollama: dengcao/Qwen3-Embedding-4B:Q4_K_M）
DIMENSION = 2560


def _run_ov(args: list, timeout: int = 60) -> str:
    """执行 ov CLI 命令，返回 stdout"""
    try:
        result = subprocess.run(
            ["ov"] + args,
            capture_output=True,
            text=True,
            timeout=timeout,
            env={**os.environ, "PATH": "/opt/homebrew/bin:" + os.environ.get("PATH", "")}
        )
        return result.stdout
    except subprocess.TimeoutExpired:
        return ""
    except Exception as e:
        print(f"[ov error] {' '.join(args)}: {e}")
        return ""


def get_embedding(text: str) -> Optional[list]:
    """
    获取文本 embedding（直接调用阿里 dashscope API）。
    仅用于 search_points 实时查询，由 OpenViking 内部生成索引。
    """
    import requests

    api_key = "sk-2c32e974b4734377a31681bd185a1b37"
    api_base = "https://dashscope.aliyuncs.com/compatible-mode/v1"

    try:
        resp = requests.post(
            f"{api_base}/embeddings",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            },
            json={"model": "text-embedding-v4", "input": text, "dimensions": DIMENSION},
            timeout=30
        )
        if resp.status_code == 200:
            return resp.json()["data"][0]["embedding"]
        print(f"[Embedding error] {resp.status_code}: {resp.text[:200]}")
    except Exception as e:
        print(f"[Embedding failed] {e}")
    return None


def search_points(query: str, top_k: int = 10, module_filter: str = None) -> list:
    """
    搜索产品知识库
    返回: [(payload_dict, score), ...]
    payload_dict 包含: text, uri, abstract, level, module 等
    """
    # ov search 返回格式化文本，需要解析
    output = _run_ov(["search", "--uri", PRODUCT_RESOURCE_URI, "-n", str(top_k), query], timeout=30)
    if not output:
        return []

    results = []
    for line in output.split("\n"):
        if line.startswith("resource") or line.startswith("memory") or line.startswith("skill"):
            parts = line.split()
            if len(parts) < 6:
                continue

            # 格式: resource uri level score abstract...
            context_type = parts[0]
            uri = parts[1]
            level = parts[2]
            score = parts[3]
            abstract = " ".join(parts[4:])

            try:
                score = float(score)
            except ValueError:
                score = 0.0

            # 跳过非 product 资源的结果
            if not uri.startswith(PRODUCT_RESOURCE_URI):
                continue

            # 从 URI 中提取文件名作为模块标识
            filename = os.path.basename(uri).replace(".md", "")

            payload = {
                "text": abstract[:500],  # 使用 abstract 作为文本内容
                "uri": uri,
                "abstract": abstract,
                "level": int(level) if level.isdigit() else 2,
                "source": "product_knowledge",
                "module": _guess_module_from_uri(uri),
                "doc_name": filename,
            }
            results.append((payload, score))

    return results


def _guess_module_from_uri(uri: str) -> str:
    """从 URI 路径猜测模块名"""
    # URI 中可能包含模块名信息（虽然有乱码，但可以从路径推断）
    uri_lower = uri.lower()
    if "gysgl" in uri_lower or "gongyings" in uri_lower:
        return "供应商管理"
    if "xunyu" in uri_lower or "xys" in uri_lower:
        return "寻源管理"
    if "cgdd" in uri_lower or "dingdan" in uri_lower:
        return "采购订单"
    if "hetong" in uri_lower or "ht" in uri_lower:
        return "合同与主数据"
    if "caiwu" in uri_lower or "cw" in uri_lower:
        return "财务结算"
    if "shuj" in uri_lower or "sj" in uri_lower:
        return "数据应用"
    if "zhineng" in uri_lower or "zn" in uri_lower:
        return "智能应用"
    return "产品功能"


def collection_exists() -> bool:
    """检查 collection 是否存在（OpenViking 自动管理）"""
    return True


def create_collection() -> bool:
    """创建 collection（OpenViking 自动创建，无需手动操作）"""
    print("[OpenViking] Collection 由 ov add-resource 自动创建，无需手动创建。")
    return True


def add_points_batch(items: list[dict], batch_size: int = 50) -> int:
    """
    批量添加记录。
    注意：产品知识已通过 ov add-resource 导入，此函数仅作兼容保留。
    如需重新导入，运行: ov add-resource <path> --to viking://resources/product
    """
    print(f"[Warning] add_points_batch 已弃用，产品知识已通过 ov add-resource 导入。")
    print(f"          如需重新导入，请运行: ov add-resource <path> --to {PRODUCT_RESOURCE_URI}")
    return len(items)


def count_points() -> int:
    """获取向量总数"""
    output = _run_ov(["observer", "vikingdb"], timeout=10)
    if not output:
        return 0

    for line in output.split("\n"):
        if "context" in line and "TOTAL" not in line and "Collection" not in line and "+" not in line:
            # 格式: |  context   |      1      |     1240     |   OK   |
            # 提取 | 之间的数字
            import re
            numbers = re.findall(r'\d+', line)
            if len(numbers) >= 2:
                return int(numbers[1])  # 第二个数字是 Vector Count (1240)
            if len(numbers) >= 1:
                return int(numbers[0])
    return 0


def delete_collection() -> bool:
    """删除 collection（通过 ov rm -r）"""
    output = _run_ov(["rm", "-r", PRODUCT_RESOURCE_URI], timeout=30)
    return "Removed" in output or "removed" in output.lower()


# ── 兼容性别名 ──────────────────────────────────────────────
COLLECTION_NAME = "product_knowledge (viking://resources/product)"
