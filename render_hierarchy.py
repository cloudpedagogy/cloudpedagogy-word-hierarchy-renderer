#!/usr/bin/env python3
"""Create a self-contained interactive hierarchy from tables in a Word document."""

from __future__ import annotations

import argparse
import html
import json
import re
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from docx import Document

VERSION = "1.0.0"
LAYOUTS = {"tree", "radial-tree", "sunburst", "icicle", "treemap", "circle-pack"}
LAYOUT_ALIASES = {
    "tidy tree": "tree", "tidy-tree": "tree", "horizontal tree": "tree",
    "radial": "radial-tree", "radial tree": "radial-tree",
    "circle packing": "circle-pack", "circle-packing": "circle-pack", "pack": "circle-pack",
}
NODE_ALIASES = {
    "id": {"id", "node id", "node_id", "key"},
    "parent_id": {"parent id", "parent_id", "parent", "parent node", "parent key"},
    "label": {"label", "name", "title", "node"},
    "type": {"type", "level", "kind", "category"},
    "value": {"value", "size", "weight", "count", "students"},
    "status": {"status", "state"},
    "description": {"description", "details", "summary", "notes"},
    "color": {"color", "colour", "fill"},
    "link": {"link", "url", "web link"},
}
SETTING_ALIASES = {
    "title": {"title", "visualisation title", "visualization title"},
    "subtitle": {"subtitle", "description"},
    "default_layout": {"default layout", "default_layout", "layout"},
    "allow_layout_switching": {"allow layout switching", "layout switching", "switch layouts"},
    "colour_by": {"colour by", "color by", "colour_by", "color_by"},
    "size_by": {"size by", "size_by"},
    "show_labels": {"show labels", "labels"},
    "theme": {"theme"},
}


@dataclass
class Issue:
    level: str
    message: str
    row: int | None = None


@dataclass
class ParseResult:
    settings: dict[str, str] = field(default_factory=dict)
    nodes: list[dict[str, Any]] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    root_id: str | None = None


def normalise(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("_", " ").strip().lower())


def clean_cell(cell) -> str:
    return "\n".join(p.text.strip() for p in cell.paragraphs if p.text.strip()).strip()


def canonical(raw: str, aliases: dict[str, set[str]]) -> str | None:
    value = normalise(raw)
    for key, variants in aliases.items():
        if value == normalise(key) or value in variants:
            return key
    return None


def valid_link(value: str) -> bool:
    parsed = urlparse(value)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def valid_color(value: str) -> bool:
    return bool(re.fullmatch(r"#[0-9a-fA-F]{3,8}", value) or re.fullmatch(r"[a-zA-Z]+", value))


def parse_number(value: str) -> float | None:
    text = value.strip().replace(",", "")
    if not text:
        return None
    try:
        number = float(text)
        return number if number >= 0 else None
    except ValueError:
        return None


def iter_blocks(document):
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    for child in document.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, document)
        elif child.tag.endswith("}tbl"):
            yield Table(child, document)


def table_rows(table) -> list[list[str]]:
    return [[clean_cell(cell) for cell in row.cells] for row in table.rows]


def detect_table(rows: list[list[str]], heading: str) -> str | None:
    title = normalise(heading)
    if title in {"nodes", "hierarchy nodes", "node data"}:
        return "nodes"
    if title in {"settings", "hierarchy settings", "configuration"}:
        return "settings"
    if not rows:
        return None
    headers = {normalise(v) for v in rows[0]}
    if headers & NODE_ALIASES["id"] and headers & NODE_ALIASES["label"]:
        return "nodes"
    if {"setting", "value"} <= headers or {"key", "value"} <= headers:
        return "settings"
    return None


def parse_settings(rows: list[list[str]], result: ParseResult) -> None:
    if len(rows) < 2:
        return
    headers = [normalise(x) for x in rows[0]]
    try:
        key_i = next(i for i, h in enumerate(headers) if h in {"setting", "key", "option", "name"})
        value_i = next(i for i, h in enumerate(headers) if h in {"value", "setting value"})
    except StopIteration:
        result.issues.append(Issue("warning", "SETTINGS table ignored: it needs Setting and Value columns."))
        return
    for row_num, row in enumerate(rows[1:], 2):
        if max(key_i, value_i) >= len(row) or not row[key_i].strip():
            continue
        key = canonical(row[key_i], SETTING_ALIASES)
        if key:
            result.settings[key] = row[value_i].strip()
        else:
            result.issues.append(Issue("warning", f"Unknown setting '{row[key_i]}' ignored.", row_num))


def parse_nodes(rows: list[list[str]], result: ParseResult) -> None:
    if len(rows) < 2:
        result.issues.append(Issue("warning", "A NODES table has no data rows."))
        return
    headers: list[str | None] = []
    seen: set[str] = set()
    for raw in rows[0]:
        key = canonical(raw, NODE_ALIASES)
        if key and key in seen:
            result.issues.append(Issue("warning", f"Duplicate '{key}' column; later column ignored."))
            key = None
        if key:
            seen.add(key)
        elif raw.strip():
            result.issues.append(Issue("warning", f"Unknown NODES column '{raw}' ignored."))
        headers.append(key)
    if not {"id", "label"} <= seen:
        result.issues.append(Issue("error", "NODES table requires ID and Label/Name columns."))
        return

    known = {node["id"] for node in result.nodes}
    for row_num, row in enumerate(rows[1:], 2):
        record = {headers[i]: value.strip() for i, value in enumerate(row) if i < len(headers) and headers[i]}
        if not any(record.values()):
            continue
        node_id, label = record.get("id", ""), record.get("label", "")
        if not node_id or not label:
            result.issues.append(Issue("error", "Missing node ID or label. Row skipped.", row_num))
            continue
        if node_id in known:
            result.issues.append(Issue("error", f"Duplicate node ID '{node_id}'. Row skipped.", row_num))
            continue
        known.add(node_id)
        raw_value = record.get("value", "")
        value = parse_number(raw_value)
        if raw_value and value is None:
            result.issues.append(Issue("warning", f"Invalid non-negative value '{raw_value}'; using 1.", row_num))
        color = record.get("color", "")
        if color and not valid_color(color):
            result.issues.append(Issue("warning", f"Invalid colour '{color}' ignored.", row_num))
            color = ""
        link = record.get("link", "")
        if link and not valid_link(link):
            result.issues.append(Issue("warning", f"Unsafe or invalid link '{link}' ignored.", row_num))
            link = ""
        result.nodes.append({
            "id": node_id, "parent_id": record.get("parent_id", ""), "label": label,
            "type": record.get("type", "Unspecified") or "Unspecified",
            "value": 1 if value is None else value,
            "status": record.get("status", ""), "description": record.get("description", ""),
            "color": color, "link": link, "_row": row_num,
        })


def validate_structure(result: ParseResult) -> None:
    by_id = {n["id"]: n for n in result.nodes}
    valid_nodes = []
    for node in result.nodes:
        parent = node["parent_id"]
        if parent and parent not in by_id:
            result.issues.append(Issue("error", f"Parent '{parent}' for node '{node['id']}' does not exist. Node removed.", node["_row"]))
        elif parent == node["id"]:
            result.issues.append(Issue("error", f"Node '{node['id']}' cannot be its own parent. Node removed.", node["_row"]))
        else:
            valid_nodes.append(node)
    result.nodes = valid_nodes
    by_id = {n["id"]: n for n in result.nodes}

    cyclic: set[str] = set()
    for node in result.nodes:
        trail: list[str] = []
        current = node["id"]
        while current and current in by_id:
            if current in trail:
                cyclic.update(trail[trail.index(current):])
                break
            trail.append(current)
            current = by_id[current]["parent_id"]
    if cyclic:
        result.issues.append(Issue("error", f"Circular parent relationship detected: {', '.join(sorted(cyclic))}. Cyclic nodes removed."))
        result.nodes = [n for n in result.nodes if n["id"] not in cyclic]

    present = {n["id"] for n in result.nodes}
    orphans = [n for n in result.nodes if n["parent_id"] and n["parent_id"] not in present]
    for node in orphans:
        result.issues.append(Issue("error", f"Node '{node['id']}' became orphaned and was removed.", node["_row"]))
    result.nodes = [n for n in result.nodes if n not in orphans]

    roots = [n for n in result.nodes if not n["parent_id"]]
    if len(roots) == 1:
        result.root_id = roots[0]["id"]
    elif len(roots) > 1:
        root_id = "__wordviz_root__"
        while any(n["id"] == root_id for n in result.nodes):
            root_id += "_"
        result.nodes.insert(0, {
            "id": root_id, "parent_id": "", "label": result.settings.get("title", "Hierarchy"),
            "type": "Root", "value": 0, "status": "", "description": "Automatically created root.",
            "color": "", "link": "", "_row": None, "_synthetic": True,
        })
        for node in roots:
            node["parent_id"] = root_id
        result.root_id = root_id
        result.issues.append(Issue("warning", f"{len(roots)} roots found; a common display root was added."))
    else:
        result.issues.append(Issue("error", "No valid root node was found."))


def read_docx(path: Path) -> ParseResult:
    result = ParseResult()
    document = Document(path)
    heading = ""
    node_tables = 0
    from docx.table import Table
    for block in iter_blocks(document):
        if isinstance(block, Table):
            rows = table_rows(block)
            kind = detect_table(rows, heading)
            if kind == "settings":
                parse_settings(rows, result)
            elif kind == "nodes":
                node_tables += 1
                parse_nodes(rows, result)
        elif block.text.strip():
            heading = block.text.strip()
    if not node_tables:
        result.issues.append(Issue("error", "No NODES table found. Add a table with ID, Parent ID and Label columns."))
    if result.nodes:
        validate_structure(result)
    if not result.nodes:
        result.issues.append(Issue("error", "No valid nodes were found."))
    return result


def bool_setting(value: str, default: bool = True) -> bool:
    if not value:
        return default
    return normalise(value) not in {"false", "no", "0", "off"}


def clean_layout(value: str) -> str:
    layout = normalise(value or "tree")
    layout = LAYOUT_ALIASES.get(layout, layout.replace(" ", "-"))
    return layout if layout in LAYOUTS else "tree"


def build_html(result: ParseResult, d3_source: str) -> str:
    settings = {
        "title": result.settings.get("title", "Interactive Hierarchy"),
        "subtitle": result.settings.get("subtitle", ""),
        "default_layout": clean_layout(result.settings.get("default_layout", "tree")),
        "allow_layout_switching": bool_setting(result.settings.get("allow_layout_switching", "true")),
        "colour_by": normalise(result.settings.get("colour_by", "type")),
        "size_by": normalise(result.settings.get("size_by", "value")),
        "show_labels": bool_setting(result.settings.get("show_labels", "true")),
        "theme": normalise(result.settings.get("theme", "light")),
    }
    nodes = [{k: v for k, v in n.items() if not k.startswith("_")} for n in result.nodes]
    data_json = json.dumps({"settings": settings, "nodes": nodes}, ensure_ascii=False).replace("</", "<\\/")
    title = html.escape(settings["title"])
    subtitle = html.escape(settings["subtitle"])
    layout_options = "".join(
        f'<option value="{x}">{label}</option>' for x, label in [
            ("tree", "Tidy tree"), ("radial-tree", "Radial tree"), ("sunburst", "Sunburst"),
            ("icicle", "Icicle"), ("treemap", "Treemap"), ("circle-pack", "Circle packing")
        ]
    )
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title><style>
:root{{--ink:#182230;--muted:#617083;--line:#d9e1ea;--panel:#fff;--bg:#f4f7fa;--accent:#2457a7}}
*{{box-sizing:border-box}}body{{margin:0;font:15px/1.45 system-ui,-apple-system,Segoe UI,sans-serif;color:var(--ink);background:var(--bg)}}
header{{padding:22px 28px 14px;background:var(--panel);border-bottom:1px solid var(--line)}}h1{{margin:0;font-size:25px}}header p{{margin:5px 0 0;color:var(--muted)}}
.controls{{display:flex;gap:10px;flex-wrap:wrap;padding:12px 28px;background:var(--panel);border-bottom:1px solid var(--line)}}
label{{font-size:12px;font-weight:700;color:var(--muted)}}input,select,button{{height:36px;border:1px solid #bdc8d5;border-radius:7px;background:#fff;padding:0 10px;color:var(--ink)}}
button{{cursor:pointer}}.grow{{flex:1;min-width:210px}}main{{display:grid;grid-template-columns:minmax(0,1fr) 300px;gap:14px;padding:14px;height:calc(100vh - 142px)}}
#chart{{background:#fff;border:1px solid var(--line);border-radius:10px;overflow:hidden;position:relative}}#details{{background:#fff;border:1px solid var(--line);border-radius:10px;padding:17px;overflow:auto}}
#details h2{{font-size:18px;margin:0 0 8px}}.meta{{color:var(--muted);font-size:13px}}.link{{display:inline-block;margin-top:8px;color:var(--accent)}}svg{{width:100%;height:100%}}
.node{{cursor:pointer}}.node text{{font-size:12px;paint-order:stroke;stroke:white;stroke-width:3px;stroke-linejoin:round}}.linkline{{fill:none;stroke:#aab6c4;stroke-opacity:.72}}
.tip{{position:fixed;display:none;pointer-events:none;background:#17202c;color:#fff;padding:7px 9px;border-radius:6px;font-size:12px;max-width:260px;z-index:4}}
#legend{{display:flex;gap:10px;flex-wrap:wrap;padding:0 28px 10px;background:#fff;font-size:12px;color:var(--muted)}}.dot{{width:10px;height:10px;border-radius:50%;display:inline-block;margin-right:4px}}
.empty{{display:grid;place-items:center;height:100%;color:var(--muted)}}@media(max-width:800px){{main{{grid-template-columns:1fr;height:auto}}#chart{{height:65vh}}#details{{min-height:160px}}}}
</style><script>{d3_source}</script></head><body>
<header><h1>{title}</h1><p>{subtitle}</p></header>
<div class="controls">
<div><label for="layout">LAYOUT</label><br><select id="layout">{layout_options}</select></div>
<div class="grow"><label for="search">SEARCH</label><br><input class="grow" id="search" placeholder="Find a node"></div>
<div><label for="type">TYPE</label><br><select id="type"><option value="">All types</option></select></div>
<div><label>&nbsp;</label><br><button id="reset">Reset view</button></div>
</div><div id="legend"></div><main><section id="chart" aria-label="Interactive hierarchy"></section>
<aside id="details"><h2>Select a node</h2><p class="meta">Click any shape to view its description and values.</p></aside></main><div class="tip"></div>
<script>
const DATA={data_json}; const chart=d3.select("#chart"), details=d3.select("#details"), tip=d3.select(".tip");
const palette=["#3569b7","#e07a32","#469d77","#8b63b7","#d05568","#aa8a2e","#318b9b","#687383"];
const types=[...new Set(DATA.nodes.map(d=>d.type||"Unspecified"))].sort();
const color=d3.scaleOrdinal(types,palette); types.forEach(t=>d3.select("#type").append("option").attr("value",t).text(t));
d3.select("#layout").property("value",DATA.settings.default_layout); if(!DATA.settings.allow_layout_switching)d3.select("#layout").attr("disabled",true);
d3.select("#legend").selectAll("span").data(types).join("span").html(d=>`<i class="dot" style="background:${{color(d)}}"></i>${{escapeHtml(d)}}`);
let transform=d3.zoomIdentity;
function escapeHtml(s){{return String(s??"").replace(/[&<>"']/g,c=>({{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}}[c]))}}
function visibleData(){{const q=d3.select("#search").property("value").trim().toLowerCase(), t=d3.select("#type").property("value"); const keep=new Set();
 const byId=new Map(DATA.nodes.map(d=>[d.id,d])); DATA.nodes.forEach(n=>{{if((!q||(n.label+" "+n.description).toLowerCase().includes(q))&&(!t||n.type===t)){{let x=n;while(x){{keep.add(x.id);x=byId.get(x.parent_id)}}}}}});
 return DATA.nodes.filter(n=>keep.has(n.id));}}
function show(d){{const n=d.data;details.html(`<h2>${{escapeHtml(n.label)}}</h2><p class="meta">${{escapeHtml(n.type)}}${{n.status?" · "+escapeHtml(n.status):""}}</p><p>${{escapeHtml(n.description)||"No description provided."}}</p><p><strong>Value:</strong> ${{n.value}}</p>${{n.link?`<a class="link" href="${{escapeHtml(n.link)}}" target="_blank" rel="noopener">Open link</a>`:""}}`)}}
function render(){{chart.selectAll("*").remove(); const data=visibleData(); if(!data.length){{chart.html('<div class="empty">No matching nodes</div>');return}}
 const root=d3.stratify().id(d=>d.id).parentId(d=>d.parent_id||null)(data); root.sum(d=>Math.max(0,+d.value||0)||1).sort((a,b)=>b.value-a.value);
 const box=chart.node().getBoundingClientRect(), w=Math.max(500,box.width), h=Math.max(420,box.height); const svg=chart.append("svg").attr("viewBox",[0,0,w,h]);
 const g=svg.append("g"); const zoom=d3.zoom().scaleExtent([.25,5]).on("zoom",e=>{{transform=e.transform;g.attr("transform",transform)}}); svg.call(zoom).call(zoom.transform,transform);
 const layout=d3.select("#layout").property("value"); if(layout==="tree"||layout==="radial-tree") drawTree(root,g,w,h,layout==="radial-tree"); else drawArea(root,g,w,h,layout);
 }}
function paint(d){{return d.data.color||color(d.data.type||"Unspecified")}}
function events(sel){{sel.on("click",(e,d)=>show(d)).on("mousemove",(e,d)=>tip.style("display","block").style("left",(e.clientX+12)+"px").style("top",(e.clientY+12)+"px").html(`<strong>${{escapeHtml(d.data.label)}}</strong><br>${{escapeHtml(d.data.type)}} · ${{d.value}}`)).on("mouseleave",()=>tip.style("display","none"))}}
function drawTree(root,g,w,h,radial){{if(radial){{const r=Math.min(w,h)/2-55;d3.tree().size([2*Math.PI,r])(root);g.attr("transform",`translate(${{w/2}},${{h/2}})`);
 g.selectAll("path").data(root.links()).join("path").attr("class","linkline").attr("d",d3.linkRadial().angle(d=>d.x).radius(d=>d.y));
 const n=g.selectAll("g.node").data(root.descendants()).join("g").attr("class","node").attr("transform",d=>`rotate(${{d.x*180/Math.PI-90}}) translate(${{d.y}},0)`);
 n.append("circle").attr("r",d=>5+Math.min(8,Math.sqrt(d.value))).attr("fill",paint); if(DATA.settings.show_labels)n.append("text").attr("x",d=>d.x<Math.PI?10:-10).attr("dy",".32em").attr("text-anchor",d=>d.x<Math.PI?"start":"end").attr("transform",d=>d.x>=Math.PI?"rotate(180)":null).text(d=>d.data.label);events(n);
 }}else{{d3.tree().nodeSize([35,175])(root); const nodes=root.descendants(), x0=d3.min(nodes,d=>d.x), x1=d3.max(nodes,d=>d.x);g.attr("transform",`translate(55,${{(h-(x1-x0))/2-x0}})`);
 g.selectAll("path").data(root.links()).join("path").attr("class","linkline").attr("d",d3.linkHorizontal().x(d=>d.y).y(d=>d.x));
 const n=g.selectAll("g.node").data(nodes).join("g").attr("class","node").attr("transform",d=>`translate(${{d.y}},${{d.x}})`);n.append("circle").attr("r",d=>5+Math.min(8,Math.sqrt(d.value))).attr("fill",paint);if(DATA.settings.show_labels)n.append("text").attr("x",11).attr("dy",".32em").text(d=>d.data.label);events(n);}}}}
function drawArea(root,g,w,h,layout){{let nodes,shape;if(layout==="sunburst"){{d3.partition().size([2*Math.PI,Math.min(w,h)/2-8])(root);g.attr("transform",`translate(${{w/2}},${{h/2}})`);shape=d3.arc().startAngle(d=>d.x0).endAngle(d=>d.x1).padAngle(.003).innerRadius(d=>d.y0).outerRadius(d=>d.y1-1);nodes=root.descendants().filter(d=>d.depth);}}
 else if(layout==="icicle"){{d3.partition().size([w,h])(root);shape=d=>`M${{d.x0}},${{d.y0}}H${{d.x1}}V${{d.y1}}H${{d.x0}}Z`;nodes=root.descendants();}}
 else if(layout==="treemap"){{d3.treemap().size([w,h]).padding(2)(root);shape=d=>`M${{d.x0}},${{d.y0}}H${{d.x1}}V${{d.y1}}H${{d.x0}}Z`;nodes=root.descendants().filter(d=>d.depth);}}
 else{{d3.pack().size([w,h]).padding(5)(root);nodes=root.descendants();shape=null}}
 const n=g.selectAll("g.node").data(nodes).join("g").attr("class","node");if(layout==="circle-pack")n.attr("transform",d=>`translate(${{d.x}},${{d.y}})`).append("circle").attr("r",d=>d.r).attr("fill",paint).attr("fill-opacity",d=>d.children?.18:.78).attr("stroke","#fff");else n.append("path").attr("d",shape).attr("fill",paint).attr("fill-opacity",d=>d.children?.62:.9).attr("stroke","#fff");
 if(DATA.settings.show_labels)n.filter(d=>layout==="circle-pack"?d.r>24:layout==="sunburst"?false:((d.x1-d.x0)>65&&(d.y1-d.y0)>20)).append("text").attr("x",d=>layout==="circle-pack"?0:d.x0+5).attr("y",d=>layout==="circle-pack"?4:d.y0+16).attr("text-anchor",d=>layout==="circle-pack"?"middle":"start").text(d=>d.data.label.length>24?d.data.label.slice(0,22)+"…":d.data.label);events(n)}}
d3.selectAll("#layout,#type").on("change",()=>{{transform=d3.zoomIdentity;render()}});d3.select("#search").on("input",render);d3.select("#reset").on("click",()=>{{transform=d3.zoomIdentity;d3.select("#search").property("value","");d3.select("#type").property("value","");render()}});addEventListener("resize",render);render();
</script></body></html>"""


def write_outputs(result: ParseResult, output: Path, d3_path: Path) -> None:
    output.mkdir(parents=True, exist_ok=True)
    d3_source = d3_path.read_text(encoding="utf-8")
    clean_nodes = [{k: v for k, v in n.items() if not k.startswith("_")} for n in result.nodes]
    (output / "data.json").write_text(json.dumps({"settings": result.settings, "nodes": clean_nodes}, indent=2, ensure_ascii=False), encoding="utf-8")
    (output / "index.html").write_text(build_html(result, d3_source), encoding="utf-8")
    lines = ["# Hierarchy QA report", "", f"- Nodes accepted: {len(result.nodes)}", f"- Root: {result.root_id or 'none'}",
             f"- Errors: {sum(i.level == 'error' for i in result.issues)}", f"- Warnings: {sum(i.level == 'warning' for i in result.issues)}", "", "## Issues", ""]
    lines += [f"- **{i.level.upper()}**{f' (row {i.row})' if i.row else ''}: {i.message}" for i in result.issues] or ["- None"]
    (output / "qa_report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Word .docx containing SETTINGS and NODES tables")
    parser.add_argument("-o", "--output", type=Path, default=Path("output/hierarchy"))
    parser.add_argument("--d3", type=Path, default=Path(__file__).parent / "vendor/d3.v7.min.js")
    parser.add_argument("--strict", action="store_true", help="Return a failure code if any errors are reported")
    parser.add_argument("--overwrite-output", action="store_true", help="Replace an existing output directory")
    parser.add_argument("--version", action="version", version=VERSION)
    args = parser.parse_args(argv)
    if not args.input.is_file() or args.input.suffix.lower() != ".docx":
        parser.error("input must be an existing .docx file")
    if not args.d3.is_file():
        parser.error(f"D3 asset not found: {args.d3}")
    if args.output.exists() and args.overwrite_output:
        shutil.rmtree(args.output)
    result = read_docx(args.input)
    write_outputs(result, args.output, args.d3)
    errors = sum(i.level == "error" for i in result.issues)
    warnings = sum(i.level == "warning" for i in result.issues)
    print(f"Created {args.output / 'index.html'} with {len(result.nodes)} nodes ({errors} errors, {warnings} warnings).")
    return 2 if (args.strict and errors) or not result.nodes else 0


if __name__ == "__main__":
    sys.exit(main())
