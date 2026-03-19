#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Qdrant operations for product-expert - 独立模块
操作 product_knowledge collection，与 Mem0 的 openclaw_memories 完全隔离
"""

import os
import sys
import time
import json
import requests
import uuid
from pathlib import Path

# API配置
QDRANT_HOST = os.environ.get("QDRANT_HOST", "localhost")
QDRANT_PORT = int(os.environ.get("QDRANT_PORT", "6333"))
COLLECTION_NAME = "product_knowledge"
DIMENSION = 3072

OPENAI_API_KEY = "sk-Pzt8a346e78b733bfead64b269317c033e97cd59abfWoqEt"
OPENAI_BASE_URL = "https://api.gptsapi.net/v1"
EMBEDDING_MODEL = "text-embedding-3-large"


def get_embedding(text: str) -> list | None:
    url = f"{OPENAI_BASE_URL}/embeddings"
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {"model": EMBEDDING_MODEL, "input": text}
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
        if resp.status_code == 200:
            return resp.json()["data"][0]["embedding"]
        print(f"Embedding error: {resp.status_code} - {resp.text}")
        return None
    except Exception as e:
        print(f"Embedding failed: {e}")
        return None


def get_client():
    from qdrant_client import QdrantClient
    return QdrantClient(host=QDRANT_HOST, port=QDRANT_PORT)


def collection_exists() -> bool:
    cols = get_client().get_collections().collections
    return COLLECTION_NAME in [c.name for c in cols]


def create_collection() -> bool:
    client = get_client()
    if collection_exists():
        print(f"Collection '{COLLECTION_NAME}' already exists.")
        return True
    try:
        from qdrant_client.models import Distance, VectorParams
        client.recreate_collection(
            collection_name=COLLECTION_NAME,
            vectors_config=VectorParams(size=DIMENSION, distance=Distance.COSINE)
        )
        print(f"Collection '{COLLECTION_NAME}' created (dim={DIMENSION}).")
        return True
    except Exception as e:
        print(f"Create failed: {e}")
        return False


def add_points_batch(items: list[dict], batch_size: int = 50) -> int:
    """
    批量添加记录
    items: [{"text": "...", "metadata": {...}}, ...]
    返回成功数
    """
    texts = [item["text"] for item in items]
    metadatas = [item["metadata"] for item in items]

    # 批量获取embedding（OpenAI API支持batch）
    url = f"{OPENAI_BASE_URL}/embeddings"
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }

    total_added = 0
    for i in range(0, len(texts), batch_size):
        batch_texts = texts[i:i + batch_size]
        batch_metas = metadatas[i:i + batch_size]

        payload = {"model": EMBEDDING_MODEL, "input": batch_texts}
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=120)
            if resp.status_code != 200:
                print(f"Batch embedding error: {resp.status_code}")
                continue
            embeddings = resp.json()["data"]
            # 按input_index排序
            embeddings.sort(key=lambda x: x["index"])
            emb_list = [e["embedding"] for e in embeddings]
        except Exception as e:
            print(f"Batch embedding failed: {e}")
            continue

        # 构建points
        from qdrant_client.models import PointStruct
        points = [
            PointStruct(
                id=str(uuid.uuid4()),
                vector=emb,
                payload={"text": text, **meta, "created_at": time.time()}
            )
            for text, meta, emb in zip(batch_texts, batch_metas, emb_list)
        ]

        try:
            get_client().upsert(collection_name=COLLECTION_NAME, points=points)
            total_added += len(points)
        except Exception as e:
            print(f"Batch upsert failed: {e}")

    return total_added


def search_points(query: str, top_k: int = 10, module_filter: str = None) -> list:
    emb = get_embedding(query)
    if emb is None:
        return []
    kwargs = {
        "collection_name": COLLECTION_NAME,
        "query_vector": emb,
        "limit": top_k,
        "with_payload": True,
    }
    if module_filter:
        from qdrant_client.models import Filter, FieldCondition, MatchKeyword
        query_filter = Filter(
            must=[FieldCondition(key="module", match=MatchKeyword(value=module_filter))]
        )
    else:
        query_filter = None
    try:
        resp = get_client().query_points(
            collection_name=COLLECTION_NAME,
            query=emb,
            query_filter=query_filter,
            limit=top_k,
            with_payload=True,
        )
        points = resp.points if hasattr(resp, 'points') else resp
        return [(p.payload, p.score) for p in points]
    except Exception as e:
        print(f"Search failed: {e}")
        return []


def count_points() -> int:
    try:
        info = get_client().get_collection(collection_name=COLLECTION_NAME)
        return info.points_count
    except:
        return 0


def delete_collection() -> bool:
    try:
        get_client().delete_collection(collection_name=COLLECTION_NAME)
        print(f"Deleted '{COLLECTION_NAME}'.")
        return True
    except Exception as e:
        print(f"Delete failed: {e}")
        return False
