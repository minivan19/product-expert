# -*- coding: utf-8 -*-
"""
三相方案工作单分析 - 可独立测试
"""
import requests, re, json, time, os, openpyxl

LLM_API_KEY = "sk-340ed7819c2346508c0a46a80df85999"
LLM_BASE_URL = "https://api.deepseek.com/v1"
LLM_MODEL = "deepseek-chat"
HIERARCHY_PATH = r"C:\Users\mingh\.openclaw\workspace\skills\product-expert\references\product_modules_hierarchy.json"

def call_llm(prompt_text, timeout=600):
    resp = requests.post(
        LLM_BASE_URL + "/chat/completions",
        headers={"Authorization": "Bearer " + LLM_API_KEY, "Content-Type": "application/json"},
        json={"model": LLM_MODEL, "messages": [{"role": "user", "content": prompt_text}],
              "temperature": 0.1, "max_tokens": 4000},
        timeout=timeout
    )
    if resp.status_code != 200:
        raise Exception("LLM错误 %s: %s" % (resp.status_code, resp.text[:200]))
    return resp.json()["choices"][0]["message"]["content"]

def load_hierarchy():
    with open(HIERARCHY_PATH, encoding="utf-8") as f:
        return json.load(f)

def build_block(records, max_desc=200):
    items = []
    for j, rec in enumerate(records, 1):
        items.append("[工单%d] 标题：%s\n描述：%s" % (
            j,
            rec.get("标题", ""),
            (rec.get("描述") or "")[:max_desc]
        ))
    return "\n\n".join(items)

def extract_category(text, key):
    """从LLM输出中提取分类列表"""
    # 尝试多种冒号形式（全角/半角）
    candidates = [
        "【%s】：" % key,
        "【%s】:" % key,
        "【%s】： " % key,
        "[%s]：" % key,
        "[%s]:" % key,
    ]
    for prefix in candidates:
        idx = text.rfind(prefix)  # rfind 更安全
        if idx == -1:
            continue
        segment = text[idx + len(prefix):]
        # 截取到下一个【或##为止
        end = len(segment)
        for p in ["【", "##", "【"]:
            e = segment.find(p)
            if e != -1 and e < end:
                end = e
        raw = segment[:end].strip()
        # 按中英文分隔符分割
        items = re.split(r'[、，,;；\n]+', raw)
        result = []
        for it in items:
            it = it.strip().strip('，。、.').strip()
            if it and len(it) > 1:
                result.append(it)
        if result:
            return result
    return []

def run(client_dir, year="2025"):
    """主入口：返回工作单分析结果"""
    hierarchy = load_hierarchy()

    # 读取工作单
    ops_dir = os.path.join(client_dir, "运维工单")
    all_records = []
    for fname in sorted(os.listdir(ops_dir)):
        if not fname.endswith(".xlsx") or year not in fname:
            continue
        fpath = os.path.join(ops_dir, fname)
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        ws = wb.active
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        title_col = desc_col = None
        for i, h in enumerate(headers):
            if h and "标题" in str(h): title_col = i
            if h and "描述" in str(h): desc_col = i
        for row in ws.iter_rows(min_row=2, values_only=True):
            title = str(row[title_col]) if title_col is not None and row[title_col] else ""
            desc = str(row[desc_col]) if desc_col is not None and row[desc_col] else ""
            if title or desc:
                all_records.append({"标题": title, "描述": desc})
        wb.close()
    print("[工作单] %d 条" % len(all_records))

    # 读取合同模块（主数据）
    bought_modules = []
    for subdir in ["基础数据", "客户档案", "其他文档"]:
        sub = os.path.join(client_dir, subdir)
        if not os.path.isdir(sub): continue
        for fn in os.listdir(sub):
            if "客户主数据" in fn and fn.endswith(".xlsx"):
                fpath = os.path.join(sub, fn)
                try:
                    wb = openpyxl.load_workbook(fpath, data_only=True)
                    ws = wb.active
                    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
                    for col_i, h in enumerate(headers):
                        if h and ("订阅模块" in str(h) or "运营模块" in str(h)):
                            for row in ws.iter_rows(min_row=2, values_only=True):
                                val = row[col_i] if col_i < len(row) else None
                                if val:
                                    parts = re.split(r'[,，、;\n]', str(val).strip())
                                    for p in parts:
                                        p = p.strip()
                                        if p and len(p) > 1:
                                            bought_modules.append(p)
                            break
                    wb.close()
                except Exception as e:
                    print("    xlsx读取失败: %s" % e)
    seen = set()
    bought_unique = []
    for m in bought_modules:
        if m not in seen:
            seen.add(m); bought_unique.append(m)
    bought_modules = bought_unique

    # Fallback：如果合同模块为空，用产品标准列表
    if not bought_modules:
        std_modules = [item["module"] for item in hierarchy]
        print("[合同模块] 无数据，用标准列表(%d个)" % len(std_modules))
        bought_modules = std_modules
    else:
        print("[合同模块] %d 个: %s" % (len(bought_modules), bought_modules[:5]))

    # 构建模块目录
    std_names = [item["module"] for item in hierarchy]
    mod_catalog = "\n".join(["- %s" % n for n in std_names])
    block = build_block(all_records)
    t0 = time.time()

    # Phase 1：三分类
    print("\n=== Phase 1：三分类判断 ===")
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
    t1 = time.time()
    raw_p1 = call_llm(prompt_p1)
    print("Phase1耗时: %.1fs" % (time.time() - t1))
    used = extract_category(raw_p1, "已用")
    bought_unused = extract_category(raw_p1, "买了没用")
    unused = extract_category(raw_p1, "未用")
    print("已用(%d): %s" % (len(used), used))
    print("买了没用(%d): %s" % (len(bought_unused), bought_unused))
    print("未用(%d): %s" % (len(unused), unused))

    # Phase 2A：已用模块深挖子功能
    print("\n=== Phase 2A：已用模块深挖子功能 ===")
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
            "2. 针对每个已用模块，列出工单中实际出现的具体功能\n"
            "3. 结构如下：\n\n"
            "## 已用模块功能分析\n"
            "### <模块名>\n"
            "  【已用功能】：<工单中实际用到的具体功能>\n"
            "  【未用功能】：<该模块下工单中完全没有涉及的功能>\n"
            "...\n\n"
            "请开始分析"
        )
        t2 = time.time()
        raw_p2a = call_llm(prompt_p2a)
        print("Phase2A耗时: %.1fs" % (time.time() - t2))
        cur_mod = None
        for line in raw_p2a.split('\n'):
            line_s = line.strip()
            if not line_s: continue
            m = re.match(r'#{1,3}\s*(.+?)\s*$', line)
            if m:
                cur_mod = m.group(1).strip()
                if cur_mod not in used_funcs:
                    used_funcs[cur_mod] = {"已用": [], "未用": []}
                continue
            if not cur_mod: continue
            for p, key in [("【已用功能】：", "已用"), ("【未用功能】：", "未用")]:
                if p in line:
                    items = re.split(r'[、，,;；\n]', line.split(p)[-1])
                    for f in items:
                        f = f.strip().strip('，。、.').strip()
                        if f and len(f) > 1:
                            used_funcs[cur_mod][key].append(f)

    # Phase 2B：买了没用模块分析障碍
    print("\n=== Phase 2B：买了没用模块分析障碍 ===")
    barriers = {}
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
            "...\n\n"
            "请开始分析"
        )
        t3 = time.time()
        raw_p2b = call_llm(prompt_p2b)
        print("Phase2B耗时: %.1fs" % (time.time() - t3))
        cur_mod = None
        for line in raw_p2b.split('\n'):
            line_s = line.strip()
            if not line_s: continue
            m = re.match(r'#{1,3}\s*(.+?)\s*$', line)
            if m:
                cur_mod = m.group(1).strip()
                if cur_mod not in barriers:
                    barriers[cur_mod] = {"障碍": [], "机会": []}
                continue
            if not cur_mod: continue
            for p, key in [("【障碍】：", "障碍"), ("【原因】：", "障碍"),
                            ("【可挖掘机会】：", "机会"), ("【机会】：", "机会")]:
                if p in line:
                    items = re.split(r'[、，,;；\n]', line.split(p)[-1])
                    for f in items:
                        f = f.strip().strip('，。、.').strip()
                        if f and len(f) > 1:
                            barriers[cur_mod][key].append(f)

    total_time = time.time() - t0
    print("\n总耗时: %.1fs" % total_time)
    print("\n=== 最终结果 ===")
    print("已用模块(%d): %s" % (len(used), used))
    print("买了没用(%d): %s" % (len(bought_unused), bought_unused))
    print("未用模块(%d): %s" % (len(unused), unused))

    return {
        "已用模块": sorted(set(used)),
        "未用模块": sorted(set(unused)),
        "买了没用模块": sorted(set(bought_unused)),
        "已用功能": used_funcs,
        "未用功能": {},
        "障碍分析": barriers,
        "_total_time": total_time,
    }

if __name__ == "__main__":
    client_dir = r"C:\Users\mingh\client-data\raw\客户档案\诺斯贝尔"
    result = run(client_dir)
    # 写结果
    out_path = r"C:\Users\mingh\.openclaw\media\诺斯贝尔_工单分析_v9.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print("结果已写: %s" % out_path)
