# -*- coding: utf-8 -*-
"""
term_map.py - 术语映射反馈系统
支持：
  - LLM提取工单中的业务术语
  - 术语→(模块, 功能)映射查表
  - 新术语自动识别并请求用户确认
  - 确认后持久化到 term_feedback.json
"""

import os
import json
import re
import sys
import unicodedata

SCRIPT_DIR = os.path.dirname(__file__)
FEEDBACK_PATH = os.path.join(SCRIPT_DIR, "term_feedback.json")


# ─────────────────────────────────────────
# 初始术语映射表（基于 product_modules_hierarchy.json 扩展）
# ─────────────────────────────────────────

BUILTIN_TERM_MAP = {
    # 敏捷协同套件 / 订单协同
    '采购订单': ('基础采购协同', '订单协同'),
    '订单确认': ('基础采购协同', '订单协同'),
    '订单协同': ('基础采购协同', '订单协同'),
    '送货单': ('基础采购协同', '交货协同'),
    '送货确认': ('基础采购协同', '交货协同'),
    'ASN': ('基础采购协同', '交货协同'),
    '交货协同': ('基础采购协同', '交货协同'),
    '交货计划': ('基础采购协同', '交货协同'),
    '预测协同': ('基础采购协同', '预测协同'),
    'VMI': ('基础采购协同', '预测协同'),
    '财务协同': ('基础采购协同', '财务协同'),
    # 敏捷协同套件 / 变更协同
    'ECN': ('高级采购协同', '变更协同管理（ECN/PCN）'),
    'PCN': ('高级采购协同', '变更协同管理（ECN/PCN）'),
    '变更协同': ('高级采购协同', '变更协同管理（ECN/PCN）'),
    # 智慧寻源套件 / 询报价
    '询价': ('基础采购寻源', '询报价业务'),
    '询价单': ('基础采购寻源', '询报价业务'),
    '报价': ('基础采购寻源', '询报价业务'),
    '竞价': ('基础采购寻源', '询竞价业务'),
    '比价': ('基础采购寻源', '询竞价业务'),
    '招标': ('基础采购寻源', '招投标业务'),
    '招投标': ('基础采购寻源', '招投标业务'),
    '价格库': ('基础采购寻源', '价格库'),
    '寻源': ('基础采购寻源', '询报价业务'),
    '供应商评选': ('基础采购寻源', '询报价业务'),
    # 供应商管理套件 / 准入
    '供应商准入': ('基础供应商管理', '供应商准入'),
    '供应商注册': ('基础供应商管理', '供应商准入'),
    '资质': ('基础供应商管理', '供应商准入'),
    '准入': ('基础供应商管理', '供应商准入'),
    '供应商基础管理': ('基础供应商管理', '供应商基础管理'),
    '供应商管理': ('基础供应商管理', '供应商基础管理'),
    # 高级供应商管理 / 绩效
    '供应商绩效': ('高级供应商管理', '供应商绩效管理'),
    '绩效考核': ('高级供应商管理', '供应商绩效管理'),
    '绩效': ('高级供应商管理', '供应商绩效管理'),
    '供应商考核': ('高级供应商管理', '供应商绩效管理'),
    # 质量管理
    '质量整改': ('质量管理', '质量整改单'),
    '质量索赔': ('质量管理', '质量索赔'),
    '来料检验': ('质量管理', '质量整改单'),
    # 合同管理
    '合同创建': ('合同管理', '合同创建'),
    '合同签署': ('合同管理', '合同在线签署'),
    '合同': ('合同管理', '合同创建'),
    # 商城采购
    '商城': ('商城采购', '商城采购执行'),
    '商品': ('商城采购', '商品维护管理'),
    'MRO': ('商城采购', '商城采购执行'),
    # 基础平台服务
    '系统配置': ('基础平台服务', '系统配置'),
    '审批': ('基础平台服务', '审批管理'),
    '工作流': ('基础平台服务', '审批管理'),
    '门户': ('基础平台服务', '门户管理'),
    '基础数据': ('基础平台服务', '基础数据配置'),
    # 库存管理
    '出入库': ('库存管理（非生物资）', '出入库管理'),
    '库存': ('库存管理（非生物资）', '库存盘点'),
    '盘点': ('库存管理（非生物资）', '库存盘点'),
    # 预算管理
    '预算': ('预算管理', '支出预算控制'),
    # 移动端
    '微信': ('移动端', '微信/钉钉嵌入集成'),
    '钉钉': ('移动端', '微信/钉钉嵌入集成'),
    '移动端': ('移动端', '工作流审批待办'),
}


# ─────────────────────────────────────────
# 反馈库读写
# ─────────────────────────────────────────

def load_feedback() -> dict:
    if os.path.exists(FEEDBACK_PATH):
        with open(FEEDBACK_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'confirmed': {}, 'pending': {}, 'rejected': []}


def save_feedback(data: dict):
    with open(FEEDBACK_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────
# 术语查询
# ─────────────────────────────────────────

def lookup_term(term: str) -> tuple | None:
    term = term.strip()
    if not term or len(term) < 2:
        return None
    data = load_feedback()
    if term in data.get('confirmed', {}):
        return tuple(data['confirmed'][term])
    if term in data.get('rejected', []):
        return None
    if term in data.get('pending', {}):
        return None
    if term in BUILTIN_TERM_MAP:
        return BUILTIN_TERM_MAP[term]
    return None


def add_pending(term: str, suggestion: tuple = None, evidence: str = ''):
    data = load_feedback()
    if term not in data['pending']:
        data['pending'][term] = {
            'suggestion': list(suggestion) if suggestion else [],
            'evidence': evidence,
        }
        save_feedback(data)


def confirm_term(term: str, module: str, feature: str):
    data = load_feedback()
    data['confirmed'][term] = [module, feature]
    if term in data.get('pending', {}):
        del data['pending'][term]
    save_feedback(data)
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    print(f"  [confirmed] '{term}' -> ({module}, {feature})")


def reject_term(term: str):
    data = load_feedback()
    if term not in data['rejected']:
        data['rejected'].append(term)
    if term in data.get('pending', {}):
        del data['pending'][term]
    save_feedback(data)


# ─────────────────────────────────────────
# LLM 提取术语
# ─────────────────────────────────────────

def _load_deepseek_config() -> dict | None:
    """从 openclaw.json 读取 DeepSeek API 配置"""
    oc_path = os.path.join(os.path.expanduser('~'), '.openclaw', 'openclaw.json')
    try:
        with open(oc_path, encoding='utf-8') as f:
            data = json.load(f)
        models = data.get('models', {})
        providers = models.get('providers', {})
        for name, cfg in providers.items():
            base_url = cfg.get('baseUrl', '')
            if 'deepseek' in base_url.lower():
                return {
                    'api_key': cfg.get('apiKey', ''),
                    'base_url': base_url,
                    'model': cfg.get('models', [{}])[0].get('id', 'deepseek-chat') if cfg.get('models') else 'deepseek-chat'
                }
    except Exception:
        pass
    return None


def _llm_available() -> bool:
    cfg = _load_deepseek_config()
    if not cfg:
        return False
    try:
        from openai import OpenAI
    except ImportError:
        return False
    try:
        client = OpenAI(api_key=cfg['api_key'], base_url=cfg['base_url'])
        resp = client.chat.completions.create(
            model=cfg['model'],
            messages=[{"role": "user", "content": "hi"}],
            max_tokens=5
        )
        return True
    except Exception:
        return False


def extract_terms_via_llm(workorders: list, model: str = None) -> list:
    """
    调用 LLM 提取工单中的业务术语。
    返回: [{term: str, evidence: str}, ...]
    """
    try:
        from openai import OpenAI
    except ImportError:
        return []

    cfg = _load_deepseek_config()
    if not cfg:
        return []
    client = OpenAI(api_key=cfg['api_key'], base_url=cfg['base_url'])
    model = model or cfg['model']

    batch_size = 10
    all_terms = []
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

    for i in range(0, len(workorders), batch_size):
        batch = workorders[i:i+batch_size]
        batch_num = i // batch_size + 1
        total_batches = (len(workorders) + batch_size - 1) // batch_size
        print(f"    [LLM] batch {batch_num}/{total_batches}")

        wo_text = '\n'.join([
            f"{j+1}. 标题：{rec.get('标题','')} | 描述：{rec.get('描述','')}"
            for j, rec in enumerate(batch)
        ])

        prompt = f"""你是SRM工单分析专家。从以下工单中提取涉及的业务术语/短语（如：采购订单、送货单、供应商准入、询价单、质量整改等）。

要求：
- 只提取实质性的业务操作术语，不要提取虚词、连接词
- 每个工单提取1-3个最重要的术语
- evidence字段填写支持这个判断的工单原文片段（限30字）

工单列表：
{wo_text}

输出格式（仅返回JSON数组，不要有其他文字）：
[{{"term": "术语", "evidence": "工单原文片段"}}]"""

        try:
            resp = client.chat.completions.create(
                model=model or 'MiniMax-Accelerate',
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=1500
            )
            raw = resp.choices[0].message.content
            json_match = re.search(r'\[.*\]', raw, re.DOTALL)
            if json_match:
                items = json.loads(json_match.group())
                for item in items:
                    if 'term' in item and 'evidence' in item:
                        item['term'] = item['term'].strip()
                        if item['term']:
                            all_terms.append(item)
        except Exception as e:
            print(f"    [LLM error] {e}")

    return all_terms


def extract_terms_fallback(workorders: list) -> list:
    """
    降级方案：基于关键词提取术语（不用LLM）
    扫描工单文本中的已知业务术语
    """
    results = []
    for rec in workorders:
        text = rec.get('标题', '') + ' ' + rec.get('描述', '')
        for term in BUILTIN_TERM_MAP:
            if term in text:
                results.append({'term': term, 'evidence': text[:80]})
    return results


# ─────────────────────────────────────────
# 主分析流程
# ─────────────────────────────────────────

def analyze_workorders(workorders: list, interactive: bool = True) -> dict:
    """
    分析工单，返回每个(模块, 功能)的命中统计。
    返回: {module: {feature: count}}
    """
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    print(f"\n  [Step3] 共 {len(workorders)} 条工单")

    # 优先用LLM提取
    if _llm_available():
        terms_raw = extract_terms_via_llm(workorders)
    else:
        print("    [Step3] LLM不可用，使用关键词降级 fallback")
        terms_raw = extract_terms_fallback(workorders)

    # 映射术语 → (module, feature)
    feature_counts = {}
    pending_terms = {}

    for item in terms_raw:
        term = item.get('term', '').strip()
        evidence = item.get('evidence', '')[:100]
        if not term:
            continue
        mapped = lookup_term(term)
        if mapped:
            module, feature = mapped
            if module not in feature_counts:
                feature_counts[module] = {}
            feature_counts[module][feature] = feature_counts[module].get(feature, 0) + 1
        else:
            if term not in pending_terms:
                pending_terms[term] = evidence

    # 交互确认新术语
    if pending_terms and interactive:
        print(f"\n  [!] 发现 {len(pending_terms)} 个未映射术语，请确认：")
        for i, (term, evidence) in enumerate(pending_terms.items(), 1):
            print(f"\n  [{i}] 术语: '{term}'")
            print(f"      证据: {evidence}")
            suggestion = None
            for known, (mod, feat) in BUILTIN_TERM_MAP.items():
                if term in known or known in term:
                    suggestion = (mod, feat)
                    break
            if suggestion:
                print(f"      建议: {suggestion[0]} / {suggestion[1]}")
                confirm = input(f"      确认? (y/n/跳过) > ").strip().lower()
                if confirm == 'y':
                    confirm_term(term, *suggestion)
                    m, f = suggestion
                    if m not in feature_counts:
                        feature_counts[m] = {}
                    feature_counts[m][f] = feature_counts[m].get(f, 0) + 1
                elif confirm == 'n':
                    reject_term(term)
            else:
                confirm = input(f"      请回复 模块名,功能名 （跳过请回车）> ").strip()
                if confirm and ',' in confirm:
                    parts = [p.strip() for p in confirm.split(',', 1)]
                    if len(parts) == 2:
                        confirm_term(term, parts[0], parts[1])
                        m, f = parts[0], parts[1]
                        if m not in feature_counts:
                            feature_counts[m] = {}
                        feature_counts[m][f] = feature_counts[m].get(f, 0) + 1

    # 汇总
    print(f"\n  [Step3 结果]")
    for module in sorted(feature_counts.keys()):
        feats = feature_counts[module]
        for feat, cnt in sorted(feats.items(), key=lambda x: -x[1]):
            bar = '#' * min(cnt, 15)
            print(f"    {module} / {feat}: {cnt:3d} {bar}")

    if not feature_counts:
        print(f"    (无命中)")

    return feature_counts


if __name__ == '__main__':
    # 简单测试
    test_wos = [
        {'标题': '送货单ASN未及时确认', '描述': '供应商未及时在系统中确认送货单'},
        {'标题': '询价单无法响应', '描述': '采购员反映部分供应商无法参与询价'},
    ]
    result = analyze_workorders(test_wos, interactive=False)
    print(f"\n结果: {result}")
