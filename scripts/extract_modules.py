#!/usr/bin/env python3
"""从 Qdrant 提取所有唯一 module 值"""
import sys, os, json
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
from qdrant_ops import get_client, COLLECTION_NAME

client = get_client()
modules = set()
offset = None

while True:
    result = client.scroll(
        collection_name=COLLECTION_NAME,
        limit=200,
        offset=offset,
        with_payload=True
    )
    points = result[0]
    if not points:
        break
    for rec in points:
        module = rec.payload.get("module", "")
        if module:
            modules.add(module)
    offset = result[1] if len(result) > 1 else None
    if offset is None:
        break

output_path = os.path.join(os.path.dirname(__file__), '..', 'references', 'modules_from_qdrant.json')
modules_list = sorted(modules)
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump({'modules': modules_list, 'count': len(modules_list)}, f, ensure_ascii=False, indent=2)
print(f"Written {len(modules_list)} modules to {output_path}")
