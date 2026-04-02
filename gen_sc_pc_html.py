#!/usr/bin/env python3
"""SC → PC → 产品套件 全景图 HTML 生成器"""
import json
from pathlib import Path

SKILL_DIR = Path(__file__).parent
OUT = SKILL_DIR / "SC_PC_全景图_网页版.html"
IDX = SKILL_DIR / "framework/product_card_index.json"
CARDS = SKILL_DIR / "framework/product_solution_cards"

SUITE_NORM = {
    "S4 - 敏捷协同套件":"协同套件","S3 - 寻源套件":"寻源套件",
    "S2 - 供应商管理套件":"供应商管理套件","S7 - 数据应用":"数据应用套件",
    "数据应用":"数据应用套件","数智化套件":"数据应用套件","大数据平台":"数据应用套件",
    "协同套件":"协同套件","供应商管理套件":"供应商管理套件",
    "质量套件":"质量套件","主数据":"主数据","BSMB套件":"BSMB套件",
}

# 加载
with open(IDX) as f:
    idx = json.load(f)
pc_raw = {}
for fp in sorted(CARDS.glob("PC_*.json")):
    try:
        d = json.load(open(fp))
        pid = d["id"]
        suites = list(dict.fromkeys(
            SUITE_NORM.get(c.get("套件","").strip(), c.get("套件",""))
            for c in d.get("产品功能组合", []) if c.get("套件","").strip()
        )) or ["未知"]
        pc_raw[pid] = dict(id=pid, name=d.get("名称",pid), suites=suites, _raw=d)
    except: pass

subdoms = []
for dm in idx["索引"]:
    for sd in dm.get("子域",[]):
        cards = [c for c in sd.get("已有卡片",[]) if c in pc_raw]
        subdoms.append(dict(id=sd["子域ID"], name=sd["子域名称"], dm=dm["域名称"], pcs=cards))

all_sx = list(dict.fromkeys(
    s for sd in subdoms for pid in sd["pcs"] for s in pc_raw[pid]["suites"]
))

# JSON 序列化（确保嵌入 JS 后语法安全）
def jstr(obj):
    return json.dumps(obj, ensure_ascii=False)

# 生成 HTML
html = [
    '<!DOCTYPE html><html lang="zh"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">',
    '<title>SC → PC → 产品套件 全景图</title>',
    '<style>',
    '*{box-sizing:border-box;margin:0;padding:0}',
    'body{font-family:system-ui,-apple-system,sans-serif;background:#f4f6f8;color:#1a2332;overflow:hidden;height:100vh}',
    '.p{display:grid;grid-template-columns:270px 248px 210px;height:100vh;overflow:hidden}',
    '.c{height:100vh;overflow-y:auto;padding:10px 8px;scrollbar-width:thin;scrollbar-color:#b0bec5 transparent}',
    '.c::-webkit-scrollbar{width:4px}.c::-webkit-scrollbar-thumb{background:#b0bec5;border-radius:2px}',
    '.c+.c{border-left:2px solid #cfd8dc}',
    '.h{position:sticky;top:0;z-index:2;padding:7px 8px;border-radius:8px;font-size:11px;font-weight:700;text-align:center;margin-bottom:10px}',
    '.sh{background:#1a3a5c;color:#fff}.ph{background:#e65100;color:#fff}.xh{background:#2e7d32;color:#fff}',
    '.g{margin-bottom:14px}',
    '.ln{font-size:10px;font-weight:800;color:#1a3a5c;padding:3px 7px 2px;letter-spacing:.9px;border-left:3px solid #1a3a5c;margin-bottom:3px}',
    '.card{padding:6px 9px;margin:3px 0;border-radius:7px;border:1.5px solid #cfd8dc;background:#e3f2fd;font-size:11px;min-height:38px;display:flex;align-items:center}',
    '.card:hover{box-shadow:0 2px 8px rgba(26,58,92,.18)}',
    '.sid{font-weight:700;font-size:10px;background:#1976d2;color:#fff;padding:1px 5px;border-radius:3px;margin-right:7px;white-space:nowrap;flex-shrink:0}',
    '.sname{color:#0d47a1;font-weight:500}',
    '.pc{padding:11px 11px 9px;margin:5px 0;border-radius:10px;border:2px solid #f9a825;background:#fff8e1;cursor:pointer;transition:all .18s;position:relative}',
    '.pc:hover{transform:translateY(-2px);box-shadow:0 4px 16px rgba(249,168,37,.35);border-color:#f57f17}',
    '.pcd{position:absolute;top:-9px;left:10px;font-size:10px;font-weight:800;background:#f9a825;color:#fff;padding:1px 7px;border-radius:4px}',
    '.pct{font-size:12px;font-weight:700;color:#4e342e;margin-top:4px;line-height:1.3}',
    '.pcs{flex-wrap:wrap;gap:3px;margin-top:5px;display:flex}',
    '.st{font-size:9px;padding:1px 7px;border-radius:10px;font-weight:600}',
    '.sx{padding:8px 11px;margin:4px 0;border-radius:8px;font-size:11px;font-weight:700;color:#fff;min-height:38px;display:flex;align-items:center}',
    '#svg{position:fixed;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:1;overflow:visible}',
    '.md{display:none;position:fixed;inset:0;background:rgba(10,20,40,.6);z-index:100;justify-content:center;align-items:center;backdrop-filter:blur(3px)}',
    '.md.o{display:flex}',
    '.mc{background:#fff;border-radius:14px;width:clamp(320px,90vw,760px);max-height:85vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 20px 60px rgba(0,0,0,.3)}',
    '.mh{padding:13px 18px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #e0e0e0;flex-shrink:0}',
    '.mh h2{font-size:14px;font-weight:800}',
    '.mb{background:#f0f0f0;border:none;border-radius:50%;width:28px;height:28px;cursor:pointer;font-size:15px;display:flex;align-items:center;justify-content:center;margin-left:12px;flex-shrink:0}',
    '.mb:hover{background:#e0e0e0}',
    '.bd{padding:14px 18px;overflow-y:auto;flex:1}',
    '.s{margin-bottom:12px}',
    '.st2{font-size:10px;font-weight:700;color:#90a4ae;text-transform:uppercase;letter-spacing:.8px;margin-bottom:5px;border-bottom:1px solid #f0f0f0;padding-bottom:3px}',
    'pre{background:#f8fafc;border:1px solid #e0e0e0;border-radius:8px;padding:12px;overflow-x:auto;font-family:Consolas,monospace;font-size:10.5px;white-space:pre-wrap;word-break:break-all;max-height:55vh;line-height:1.5}',
    '</style></head><body>',
    '<div class="p">',
    '<div class="c sh" id="sc"><div class="h">业务能力（SC / 子域）</div></div>',
    '<div class="c ph" id="pc"><div class="h">方案卡（PC）</div></div>',
    '<div class="c xh" id="sx"><div class="h">产品套件</div></div>',
    '</div>',
    '<svg id="svg"></svg>',
    '<div class="md" id="md"><div class="mc"><div class="mh"><h2 id="mt"></h2><button class="mb" id="mbBtn">&#x2715;</button></div><div class="bd" id="bd"></div></div></div>',
    '<script>',
    'var D_SC='+jstr(subdoms)+';',
    'var D_PC='+jstr(pc_raw)+';',
    'var D_SX='+jstr(all_sx)+';',
    'var SS={"协同套件":["#e1f5fe","#0277bd"],"寻源套件":["#fff3e0","#e65100"],'
    '"供应商管理套件":["#f3e5f5","#6a1b9a"],"数据应用套件":["#e8f5e9","#2e7d32"],'
    '"质量套件":["#fce4ec","#ad1457"],"主数据":["#efebe9","#4e342e"],'
    '"BSMB套件":["#e0f7fa","#006064"],"未知":["#f5f5f5","#616161"]};',
    'var SC_C=["#e3f2fd","#e8f0fe","#e0f7fa","#f3e5f5","#e8f5e9","#fff8e1","#fce4ec","#efebe9","#f1f8e9"];',
    'function tag(s){var q=SS[s]||SS["未知"];return\'<span class="st" style="background:\'+q[0]+\';color:\'+q[1]+\'">\'+s+\'</span>\'}',
    'function bz(x1,y1,x2,y2){var mx=(x1+x2)/2;return"M"+x1+","+y1+" C"+mx+","+y1+" "+mx+","+y2+" "+x2+","+y2}',
    'function line(d,sw,op){var p=document.createElementNS("http://www.w3.org/2000/svg","path");'
    'p.setAttribute("d",d);p.setAttribute("stroke",sw>1.45?"#78909c":"#66bb6a");'
    'p.setAttribute("stroke-width",sw);p.setAttribute("fill","none");p.setAttribute("opacity",op);return p}',
    'function render(){',
    'var dmM={},scE=document.getElementById("sc"),pcE=document.getElementById("pc"),sxE=document.getElementById("sx");',
    'for(var i=0;i<D_SC.length;i++){var sd=D_SC[i];if(!dmM[sd.dm])dmM[sd.dm]=[];dmM[sd.dm].push(sd)}',
    'var dmK=Object.keys(dmM),dmI=0;',
    'for(var di=0;di<dmK.length;di++){var dn=dmK[di],sds=dmM[dn],g=document.createElement("div");g.className="g";',
    'var lb=document.createElement("div");lb.className="ln";lb.textContent=dn;g.appendChild(lb);',
    'for(var si=0;si<sds.length;si++){var sd2=sds[si],c=document.createElement("div");c.className="card";'
    'c.style.background=SC_C[dmI%SC_C.length];'
    'c.innerHTML=\'<span class="sid">\'+sd2.id+\'</span><span class="sname">\'+sd2.name+\'</span>\';g.appendChild(c);}',
    'scE.appendChild(g);dmI++;}',
    'var seenPC={},seenSX={};',
    'for(var i2=0;i2<D_SC.length;i2++){var sd3=D_SC[i2];',
    'for(var pi=0;pi<sd3.pcs.length;pi++){var pid=sd3.pcs[pi];',
    'if(seenPC[pid])continue;seenPC[pid]=true;',
    'var pcd=D_PC[pid],el=document.createElement("div");el.className="pc";el.dataset.p=pid;',
    'el.innerHTML=\'<div class="pcd">\'+pid+\'</div><div class="pct">\'+pcd.name+\'</div><div class="pcs"></div>\';',
    'var tg=el.querySelector(".pcs");for(var ti=0;ti<pcd.suites.length;ti++)tg.innerHTML+=tag(pcd.suites[ti]);',
    'el.onclick=(function(p){return function(){showModal(p)}})(pid);pcE.appendChild(el);}}',
    'for(var si2=0;si2<D_SX.length;si2++){var s=D_SX[si2];',
    'if(seenSX[s])continue;seenSX[s]=true;',
    'var q=SS[s]||SS["未知"],el2=document.createElement("div");el2.className="sx";',
    'el2.style.background=q[0];el2.style.color=q[1];',
    'el2.innerHTML=\'<span>\'+s+\'</span>\';sxE.appendChild(el2);}}',
    'function drawLines(){',
    'var svg=document.getElementById("svg");svg.innerHTML="";',
    'svg.setAttribute("width",window.innerWidth);svg.setAttribute("height",window.innerHeight);',
    'var scR=document.getElementById("sc").getBoundingClientRect().right,',
    'pcL=document.getElementById("pc").getBoundingClientRect().left,',
    'sxL=document.getElementById("sx").getBoundingClientRect().left;',
    'function ym(sel){var m={},es=document.querySelectorAll(sel);',
    'for(var i=0;i<es.length;i++){var r=es[i].getBoundingClientRect();',
    'm[es[i].textContent.trim().slice(0,30)]=r.top+r.height/2}return m}',
    'var scY=ym(".card"),pcY=ym(".pc"),sxY=ym(".sx");',
    'for(var i=0;i<D_SC.length;i++){var sd=D_SC[i];',
    'if(!(sd.id in scY))continue;var sy=scY[sd.id];',
    'for(var pi=0;pi<sd.pcs.length;pi++){var pid=sd.pcs[pi];',
    'if(!(pid in pcY))continue;',
    'svg.appendChild(line(bz(scR,sy,pcL,pcY[pid]),1.5,0.5));}}',
    'for(var pk in D_PC){var pcd=D_PC[pk];if(!(pk in pcY))continue;var py=pcY[pk];',
    'for(var si=0;si<pcd.suites.length;si++){var s=pcd.suites[si];',
    'if(!(s in sxY))continue;',
    'svg.appendChild(line(bz(pcL,py,sxL,sxY[s]),1.4,0.55));}}}',
    'function showModal(pid){',
    'var pcd=D_PC[pid];if(!pcd||!pcd._raw)return;',
    'var raw=pcd._raw,keys=["版本","状态","对应业务卡","业务流程说明","适用条件","实施步骤","二开介入点","关键配置项","关联方案卡","更新时间"];',
    'var html="";',
    'for(var ki=0;ki<keys.length;ki++){var k=keys[ki];',
    'if(raw[k]===undefined)continue;',
    'var v=Array.isArray(raw[k])?raw[k].join("\\n"):JSON.stringify(raw[k],null,2);',
    'html+=\'<div class="s"><div class="st2">\'+k+\'</div><pre>\'+v.replace(/</g,"&lt;")+\'</pre></div>\'}',
    'if(raw.产品功能组合&&raw.产品功能组合.length){',
    'var rows=[];for(var ri=0;ri<raw.产品功能组合.length;ri++){var c=raw.产品功能组合[ri];',
    'rows.push((c.功能||"")+" | "+(c.套件||"")+" | "+(c.角色||""))}',
    'html+=\'<div class="s"><div class="st2">产品功能组合</div><pre>\'+rows.join("\\n")+\'</pre></div>\'}',
    'document.getElementById("mt").textContent=pid+" | "+pcd.name;',
    'document.getElementById("bd").innerHTML=html;',
    'document.getElementById("md").classList.add("o")}',
    'document.getElementById("md").onclick=function(e){if(e.target===this)this.classList.remove("o")};',
    'document.getElementById("mbBtn").onclick=function(){document.getElementById("md").classList.remove("o")};',
    'document.addEventListener("keydown",function(e){if(e.key==="Escape")document.getElementById("md").classList.remove("o")});',
    'render();setTimeout(drawLines,150);window.addEventListener("resize",function(){setTimeout(drawLines,150)});',
    '</script></body></html>',
]

with open(OUT,"w",encoding="utf-8") as f:
    f.write('\n'.join(html))

size = len(open(OUT).read())
print(f"OK {OUT.name} ({size//1024}KB)")
print(f"  子域:{len(subdoms)} PC:{len(pc_raw)} 套件:{len(all_sx)}")
