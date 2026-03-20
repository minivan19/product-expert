# -*- coding: utf-8 -*-
"""
gap_analysis.py - 客户功能缺口分析
三步分析框架：
  Step 1: 买了没？    → 合同/主数据 + 关键词匹配
  Step 2: 实施了没？  → 蓝图方案 + 内容关键词匹配
  Step 3: 用了没？    → 运维工单 + LLM凝练
基准：references/product_modules_hierarchy.json
"""

import os
import re
import json
import sys
import openpyxl
import win32com.client
import pythoncom
import zipfile
import fitz

CLIENT_DATA_ROOT = r"C:\Users\mingh\client-data\raw\客户档案"
HIERARCHY_PATH = os.path.join(os.path.dirname(__file__), "..", "references", "product_modules_hierarchy.json")


# -----------------------------------------
# 工具函数
# -----------------------------------------

def norm(s: str) -> str:
    """消除全角ASCII/标点，CJK字符不变"""
    return (s
        .replace('\uff08', '(').replace('\uff09', ')')
        .replace('\uff0f', '/')
        .replace('\uff0c', ',').replace('\uff0e', '.').replace('\uff1b', ';').replace('\uff1a', ':')
    )


def load_hierarchy() -> list:
    """加载产品模块层级JSON"""
    with open(HIERARCHY_PATH, encoding='utf-8') as f:
        return json.load(f)


def build_module_kw_map(hierarchy: list) -> dict:
    """
    为每个模块构建关键词集合（用于蓝图/合同匹配）
    key = 模块名（去序号前缀）
    value = {module: str, features: set, all_keywords: set}
    """
    mod_map = {}
    for item in hierarchy:
        mod_raw = item['module']
        # 去掉序号前缀："1.基础供应商管理" → "基础供应商管理"
        if mod_raw[0].isdigit() and '.' in mod_raw[:3]:
            mod_name = mod_raw.split('.', 1)[1]
        else:
            mod_name = mod_raw
        
        kws = set()
        kws.add(mod_name)
        if mod_name not in kws:
            kws.add(mod_name)
        for feat in item.get('features', []):
            kws.add(feat)
        
        mod_map[mod_name] = {
            'module': mod_raw,
            'suite': item.get('suite', ''),
            'features': set(item.get('features', [])),
            'all_keywords': kws
        }
    return mod_map


# Excel购买模块名 → 产品模块名 的映射（已确认诺斯贝尔的命名差异）
EXCEL_TO_PRODUCT = {
    '基础供应商管理': '基础供应商管理',
    '寻源管理': '基础采购寻源',
    '基础采购协同': '基础采购协同',
    '商城采购': '商城采购',
    '自有供应商目录化': '自有供应商目录化',
    '高级供应商管理': '高级供应商管理',
    '供应商管理': '基础供应商管理',  # 部分匹配
}


# -----------------------------------------
# Step 1: 买了没？（合同优先 → 主数据备选）
# -----------------------------------------

def read_bought_from_master(client_dir: str) -> set:
    """
    从客户主数据xlsx读取"购买模块"列，返回产品模块名集合
    """
    master_path = os.path.join(client_dir, '基础数据')
    if not os.path.isdir(master_path):
        return set()
    
    xlsx_files = [f for f in os.listdir(master_path) if f.endswith('.xlsx') and not f.startswith('~$')]
    for fn in xlsx_files:
        if '客户主数据' in fn or '主数据' in fn:
            fp = os.path.join(master_path, fn)
            break
    else:
        return set()
    
    wb = openpyxl.load_workbook(fp, data_only=True)
    ws = wb.active
    row1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    col_map = {str(h).strip() if h else '': i for i, h in enumerate(row1)}
    
    # 找"购买模块"或"产品模块"列
    mod_col = None
    for name, idx in col_map.items():
        if '购买模块' in name or '产品模块' in name:
            mod_col = idx
            break
    
    if mod_col is None:
        wb.close()
        return set()
    
    bought = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[mod_col]:
            raw = str(row[mod_col]).strip()
            if raw and raw not in ('无', '空', 'None'):
                # 处理"基础供应商管理,寻源管理,..."格式
                for part in raw.split(','):
                    part = part.strip()
                    if not part:
                        continue
                    # 映射到产品模块名
                    product_mod = EXCEL_TO_PRODUCT.get(part, part)
                    bought.add(product_mod)
    
    wb.close()
    return bought


def read_bought_from_contracts(client_dir: str) -> set:
    """
    从订阅合同明细xlsx读取合同产品名，返回产品模块名集合
    合同数据是产品级，不是模块级，所以要做额外解析
    """
    sub_path = os.path.join(client_dir, '订阅合同行')
    if not os.path.isdir(sub_path):
        return set()
    
    xlsx_files = [f for f in os.listdir(sub_path) if f.endswith('.xlsx') and not f.startswith('~$')]
    found = None
    for fn in xlsx_files:
        if '明细' in fn or '订阅' in fn:
            found = os.path.join(sub_path, fn)
            break
    
    if not found:
        return set()
    
    wb = openpyxl.load_workbook(found, data_only=True)
    ws = wb.active
    row1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    col_map = {str(h).strip() if h else '': i for i, h in enumerate(row1)}
    
    # 产品名称列
    prod_col = col_map.get('产品名称', None)
    if prod_col is None:
        wb.close()
        return set()
    
    bought = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[prod_col]:
            raw = str(row[prod_col]).strip()
            if raw and raw not in ('无', '空', 'None'):
                # 合同中的产品名通常包含SRM/供应商/寻源等关键词
                # 这里根据关键词推断模块（合同级，无法精确到模块）
                if '供应商' in raw: bought.add('基础供应商管理')
                if '寻源' in raw: bought.add('基础采购寻源')
                if '采购' in raw and '协同' in raw: bought.add('基础采购协同')
                if '商城' in raw: bought.add('商城采购')
                if '目录' in raw: bought.add('自有供应商目录化')
    
    wb.close()
    return bought


def step1_bought_modules(client_dir: str) -> dict:
    """
    Step 1: 买了哪些模块？
    优先从合同读取，其次从主数据
    返回: {产品模块名: True/False}
    """
    hierarchy = load_hierarchy()
    mod_names = {m['module'].split('.', 1)[1] if '.' in m['module'][:3] else m['module']
                 for m in hierarchy}
    
    # 优先合同
    contract_bought = read_bought_from_contracts(client_dir)
    master_bought = read_bought_from_master(client_dir)
    
    # 合并：合同优先，合同没有的用主数据补充
    bought = contract_bought | master_bought
    
    result = {}
    for mod in sorted(mod_names):
        result[mod] = mod in bought
    
    print(f"  [Step1 买了没] 合同:{len(contract_bought)} 主数据:{len(master_bought)} → 合并:{len(bought)}")
    for mod, has in sorted(result.items()):
        if has:
            print(f"    [Y] {mod}")
        else:
            print(f"    [N] {mod}")
    return result


# -----------------------------------------
# Step 2: 实施了没？（蓝图方案内容匹配）
# -----------------------------------------

def extract_doc_text(fp: str) -> str:
    """读取DOC/DOCX正文"""
    text_parts = []
    word, doc = None, None
    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(os.path.abspath(fp), ReadOnly=True, ConfirmConversions=False)
        for para in doc.Paragraphs:
            t = para.Range.Text.strip()
            if t and len(t) > 1:
                text_parts.append(t)
    except Exception as e:
        text_parts.append(f"[DOC错误: {e}]")
    finally:
        if doc:
            try: doc.Close(False)
            except: pass
        if word:
            try: word.Quit()
            except: pass
        del word, doc
    pythoncom.CoUninitialize()
    return '\n'.join(text_parts)


def extract_pptx_text(fp: str) -> str:
    """读取PPTX/PPT正文（ZIP+XML方式）"""
    text_parts = []
    try:
        with zipfile.ZipFile(fp) as z:
            for name in z.namelist():
                if re.match(r'ppt/slides/slide\d+\.xml', name):
                    with z.open(name) as f:
                        content = f.read().decode('utf-8', errors='replace')
                        texts = re.findall(r'<a:t[^>]*>([^<]+)</a:t>', content)
                        for t in texts:
                            t = t.strip()
                            if t:
                                text_parts.append(t)
    except Exception as e:
        return f"[PPTX错误: {e}]"
    return '\n'.join(text_parts)


def extract_blueprint(fp: str) -> str:
    """统一入口：根据扩展名调用不同提取方法"""
    ext = os.path.splitext(fp)[1].lower()
    if ext in ('.doc', '.docx'):
        return extract_doc_text(fp)
    elif ext in ('.pptx', '.ppt'):
        return extract_pptx_text(fp)
    elif ext == '.pdf':
        # PDF流程图：正文提取失败，用文件名降级
        name = os.path.splitext(os.path.basename(fp))[0]
        return name
    return ""


def step2_implemented_modules(client_dir: str) -> dict:
    """
    Step 2: 哪些模块在蓝图中被实施？
    遍历蓝图文件，用模块/功能的关键词匹配内容
    返回: {产品模块名: {'implemented': set(), 'files': [file_list]}}
    """
    blueprint_dir = os.path.join(client_dir, '蓝图方案')
    if not os.path.isdir(blueprint_dir):
        print("  [Step2 实施了没] 蓝图文件夹不存在")
        return {}
    
    hierarchy = load_hierarchy()
    mod_kw_map = build_module_kw_map(hierarchy)
    
    # 每个模块追踪：实现了哪些功能
    impl = {mod: {'implemented': set(), 'files': set()} for mod in mod_kw_map}
    
    files = sorted(os.listdir(blueprint_dir))
    print(f"  [Step2 实施了没] 扫描蓝图文件: {len(files)}个")
    
    for fn in files:
        fp = os.path.join(blueprint_dir, fn)
        text = extract_blueprint(fp)
        text_norm = norm(text)
        
        for mod_name, mod_info in mod_kw_map.items():
            kws = mod_info['all_keywords']
            hit_feats = []
            for kw in kws:
                if norm(kw) in text_norm:
                    hit_feats.append(kw)
            if hit_feats:
                impl[mod_name]['files'].add(fn)
                for feat in hit_feats:
                    if feat in mod_info['features']:
                        impl[mod_name]['implemented'].add(feat)
    
    # 汇总输出
    for mod_name, data in sorted(impl.items()):
        total_feats = len(mod_kw_map[mod_name]['features'])
        impl_count = len(data['implemented'])
        status = f"{impl_count}/{total_feats}" if total_feats > 0 else "?"
        print(f"    {'[Y]' if data['files'] else '[N]'} {mod_name}: {status}功能, {len(data['files'])}个文件")
    
    return impl


# -----------------------------------------
# Step 3: 用了没？（运维工单 + LLM）
# -----------------------------------------

def _read_workorder_batch(fp, year):
    """读取单个工单xlsx"""
    records = []
    try:
        wb = openpyxl.load_workbook(fp, data_only=True)
        ws = wb.active
        row1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        col_map = {}
        for i, h in enumerate(row1):
            if h:
                col_map[str(h).strip()] = i
        tc = col_map.get('标题', col_map.get('工单号', 1))
        dc = col_map.get('描述', col_map.get('详细', col_map.get('问题描述', 9)))
        mod_i = col_map.get('模块', 7)
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            title = str(row[tc]).strip() if row[tc] else ''
            desc = str(row[dc]).strip() if row[dc] else ''
            mod = str(row[mod_i]).strip() if row[mod_i] else ''
            if title or desc:
                records.append({"模块": mod, "标题": title, "描述": desc[:300]})
        wb.close()
    except Exception:
        pass
    return records


def read_workorders(client_dir: str, year: int) -> list:
    """读取所有工单记录"""
    ops_dir = os.path.join(client_dir, '运维工单')
    if not os.path.isdir(ops_dir):
        return []
    
    all_records = []
    files = sorted([f for f in os.listdir(ops_dir) if f.endswith('.xlsx')])
    for fn in files:
        all_records.extend(_read_workorder_batch(os.path.join(ops_dir, fn), year))
    return all_records


def step3_used_modules(client_dir: str, year: int = 2025) -> dict:
    """
    Step 3: 工单中使用了哪些模块？（LLM凝练）
    目前先用关键词匹配（LITE模式），后续切LLM
    返回: {产品模块名: count}
    """
    hierarchy = load_hierarchy()
    mod_kw_map = build_module_kw_map(hierarchy)
    
    # 工单标签 → 产品模块 映射（从旧代码继承，已验证有效）
    WO_LABEL_MAP = {
        '订单/物流': '基础采购协同',
        '寻源（询价/招标）': '基础采购寻源',
        '系统基础/报表/应用商店': '基础平台服务',
        '合作伙伴': None,  # 无法映射
        '结算/质量': '质量管理',
        '协议（合同）/价格库': '合同管理',
        '需求（采购申请）': '基础采购寻源',
        '审批（工作流）': '基础平台服务',
    }
    
    records = read_workorders(client_dir, year)
    print(f"  [Step3 用了没] 工单记录: {len(records)}条")
    
    usage = {mod: 0 for mod in mod_kw_map}
    
    for rec in records:
        rec_text = norm(rec.get('标题', '') + ' ' + rec.get('描述', ''))
        mod_label = rec.get('模块', '')
        
        # 标签映射
        mapped_mod = WO_LABEL_MAP.get(mod_label, None)
        if mapped_mod and mapped_mod in usage:
            usage[mapped_mod] += 1
            continue
        
        # 关键词内容匹配
        for mod_name, mod_info in mod_kw_map.items():
            kws = mod_info['all_keywords']
            for kw in kws:
                if len(kw) >= 2 and norm(kw) in rec_text:
                    usage[mod_name] += 1
                    break
    
    for mod, cnt in sorted(usage.items(), key=lambda x: -x[1]):
        bar = '#' * min(cnt, 20)
        print(f"    {'[Y]' if cnt > 0 else '·'} {mod}: {cnt}次 {bar}")
    
    return usage


# -----------------------------------------
# 三步综合 → 3×2 网格分类
# -----------------------------------------

def classify_3x2(bought: dict, implemented: dict, used: dict, hierarchy: list) -> dict:
    """
    综合三步结果，输出3×2网格分类
    维度1: 买了没（bought）
    维度2: 用了多少（used count）
    
    A=买了+深度使用(>5)  B=买了+轻度使用(1-5)  C=买了+未使用(0)
    D=没买+工单有        E=没买+工单无
    """
    mod_names = [m['module'].split('.', 1)[1] if '.' in m['module'][:3] else m['module']
                 for m in hierarchy]
    
    result = {'A': [], 'B': [], 'C': [], 'D': [], 'E': []}
    
    for mod in sorted(mod_names):
        has_bought = bought.get(mod, False)
        cnt = used.get(mod, 0)
        
        if has_bought:
            if cnt > 5:
                result['A'].append(mod)
            elif cnt > 0:
                result['B'].append(mod)
            else:
                result['C'].append(mod)
        else:
            if cnt > 0:
                result['D'].append(mod)
            else:
                result['E'].append(mod)
    
    print(f"\n  [3×2分类结果]")
    labels = {'A': '深度应用(买了+高频>5)', 'B': '激活不足(买了+低频1-5)', 
               'C': '买了未实施(买了+0)', 'D': '潜在需求(没买+工单有)', 'E': '空白机会(没买+工单无)'}
    for cls in 'ABCDE':
        mods = result[cls]
        print(f"    [{cls}] {labels[cls]}: {len(mods)}个 {' '.join(mods[:3])}{'...' if len(mods)>3 else ''}")
    
    return result


# -----------------------------------------
# LLM 推荐生成（Phase A/B/C/D）
# -----------------------------------------

def call_llm(messages: list, model: str = None) -> str:
    """调用LLM接口"""
    try:
        from openai import OpenAI
    except ImportError:
        print("    [WARN] openai not installed, skipping LLM call")
        return "[LLM未安装]"
    
    client = OpenAI(api_key=os.environ.get('OPENAI_API_KEY', ''), 
                     base_url=os.environ.get('OPENAI_API_BASE', 'https://api.minimaxi.com/v1'))
    try:
        resp = client.chat.completions.create(
            model=model or 'MiniMax-Accelerate',
            messages=messages,
            temperature=0.7,
            max_tokens=2000
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[LLM错误: {e}]"


def generate_recommendations(grid: dict, used: dict, client_name: str) -> dict:
    """为各分类生成推荐内容（简化版）"""
    labels = {
        'A': '深度应用模块',
        'B': '激活不足模块',
        'C': '买了未实施模块',
        'D': '潜在需求模块',
    }
    
    results = {}
    for cls in 'ABCD':
        if not grid.get(cls):
            results[cls] = {}
            continue
        
        mod_list = ', '.join(grid[cls])
        prompt = f"""你是SRM客户成功顾问。为{client_name}的以下模块生成激活/推进建议：

{labels[cls]}：{mod_list}

请简洁回复，每模块一条建议（50字以内），格式：
模块名：建议内容
"""
        raw = call_llm([{"role": "user", "content": prompt}])
        results[cls] = {'raw': raw, 'mods': grid[cls]}
    
    return results


# -----------------------------------------
# 主入口
# -----------------------------------------

def find_client_dir(client_name: str) -> str:
    """查找客户数据目录"""
    if not os.path.isdir(CLIENT_DATA_ROOT):
        raise FileNotFoundError(f"客户档案目录不存在: {CLIENT_DATA_ROOT}")
    
    for name in os.listdir(CLIENT_DATA_ROOT):
        if client_name in name:
            return os.path.join(CLIENT_DATA_ROOT, name)
    raise FileNotFoundError(f"未找到客户目录: {client_name}")


def main(client_name: str, year: int = 2025, output_path: str = None):
    """三步分析主流程"""
    print(f"\n{'='*50}")
    print(f"客户: {client_name} | 年份: {year}")
    print(f"{'='*50}")
    
    hierarchy = load_hierarchy()
    client_dir = find_client_dir(client_name)
    
    # Step 1: 买了没
    print(f"\n[Step1] 买了哪些模块？")
    bought = step1_bought_modules(client_dir)
    
    # Step 2: 实施了没
    print(f"\n[Step2] 哪些模块已实施？")
    impl = step2_implemented_modules(client_dir)
    
    # Step 3: 用了没
    print(f"\n[Step3] 工单中使用情况？")
    used = step3_used_modules(client_dir, year)
    
    # 3×2 分类
    print(f"\n[综合] 3×2网格分类")
    grid = classify_3x2(bought, impl, used, hierarchy)
    
    # 生成推荐
    print(f"\n[推荐] 生成激活建议...")
    recs = generate_recommendations(grid, used, client_name)
    
    # 组装报告
    report = build_report(client_name, bought, impl, used, grid, recs)
    
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f"\n报告已保存: {output_path}")
    else:
        print(report)
    
    return grid, recs


def build_report(client_name, bought, impl, used, grid, recs) -> str:
    """组装Markdown报告"""
    hierarchy = load_hierarchy()
    mod_kw_map = build_module_kw_map(hierarchy)
    
    lines = [
        f"# {client_name} - 功能缺口分析与激活建议",
        f"\n**时间**: 2026-03-20",
        "\n## 执行摘要\n",
    ]
    
    for cls, label in [('A','深度应用'), ('B','激活不足'), ('C','买了未实施'), ('D','潜在需求'), ('E','空白机会')]:
        if grid.get(cls):
            lines.append(f"- **[{cls}] {label}**: {', '.join(grid[cls])}")
    
    lines.append("\n## 详细分析\n")
    
    for cls in 'ABCD':
        mods = grid.get(cls, [])
        if not mods:
            continue
        lines.append(f"### {cls}类\n")
        raw = recs.get(cls, {}).get('raw', '')
        if raw and not raw.startswith('['):
            lines.append(f"{raw}\n")
        else:
            lines.append(f"_（无详细建议）_\n")
    
    return '\n'.join(lines)


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('client', help='客户名称（部分匹配）')
    parser.add_argument('--year', type=int, default=2025)
    parser.add_argument('--output', help='输出Markdown文件路径')
    args = parser.parse_args()
    
    try:
        main(args.client, args.year, args.output)
    except FileNotFoundError as e:
        print(f"错误: {e}")
