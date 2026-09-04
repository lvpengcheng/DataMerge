from pathlib import Path
from xml.etree.ElementTree import Element, SubElement, ElementTree, register_namespace

from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "visio" / "智能组表执行流程图_单页版.vdx"
PREVIEW = ROOT / "output" / "images" / "智能组表执行流程图_单页版.png"
PDF_OUT = ROOT / "output" / "pdf" / "智能组表执行流程图_单页版.pdf"

NS = "http://schemas.microsoft.com/visio/2003/core"
VX = "http://schemas.microsoft.com/visio/2006/extension"
register_namespace("", NS)
register_namespace("vx", VX)

PAGE_W = 8.2677
PAGE_H = 11.6929


def q(tag):
    return f"{{{NS}}}{tag}"


def add_text_style(shape, font_size=0.135, bold=False, color="#24292F"):
    block = SubElement(shape, q("TextBlock"))
    SubElement(block, q("VerticalAlign")).text = "1"
    SubElement(block, q("DefaultTabStop"), {"Unit": "IN_F"}).text = "0.3937"
    char = SubElement(shape, q("Char"), {"IX": "0"})
    SubElement(char, q("Font")).text = "0"
    SubElement(char, q("Color")).text = color
    SubElement(char, q("Size"), {"Unit": "IN_F"}).text = str(font_size)
    SubElement(char, q("Style")).text = "1" if bold else "0"
    para = SubElement(shape, q("Para"), {"IX": "0"})
    SubElement(para, q("HorzAlign")).text = "1"
    SubElement(para, q("SpLine")).text = "-1"
    SubElement(para, q("SpBefore")).text = "0"
    SubElement(para, q("SpAfter")).text = "0"


def add_line_fill(shape, fill, stroke):
    line = SubElement(shape, q("Line"))
    SubElement(line, q("LineWeight"), {"Unit": "IN_F"}).text = "0.0139"
    SubElement(line, q("LineColor")).text = stroke
    SubElement(line, q("LinePattern")).text = "1"
    fill_el = SubElement(shape, q("Fill"))
    SubElement(fill_el, q("FillForegnd")).text = fill
    SubElement(fill_el, q("FillBkgnd")).text = "#FFFFFF"
    SubElement(fill_el, q("FillPattern")).text = "1"


def add_xform(shape, x, y, w, h):
    xf = SubElement(shape, q("XForm"))
    SubElement(xf, q("PinX"), {"Unit": "IN_F"}).text = str(x)
    SubElement(xf, q("PinY"), {"Unit": "IN_F"}).text = str(y)
    SubElement(xf, q("Width"), {"Unit": "IN_F"}).text = str(w)
    SubElement(xf, q("Height"), {"Unit": "IN_F"}).text = str(h)
    SubElement(xf, q("LocPinX"), {"Unit": "IN_F"}).text = str(w / 2)
    SubElement(xf, q("LocPinY"), {"Unit": "IN_F"}).text = str(h / 2)
    SubElement(xf, q("Angle")).text = "0"
    SubElement(xf, q("FlipX")).text = "0"
    SubElement(xf, q("FlipY")).text = "0"


def add_node(shapes, sid, name, label, x, y, w=2.65, h=0.48, kind="rect", fill="#F2F4F7", stroke="#98A2B3"):
    shape = SubElement(shapes, q("Shape"), {
        "ID": str(sid), "NameU": f"{name}.{sid}", "Name": label,
        "Type": "Shape", "LineStyle": "0", "FillStyle": "0", "TextStyle": "0"
    })
    add_xform(shape, x, y, w, h)
    add_line_fill(shape, fill, stroke)
    add_text_style(shape, 0.132, kind in ("start", "output"))
    geom = SubElement(shape, q("Geom"), {"IX": "0"})
    SubElement(geom, q("NoFill")).text = "0"
    SubElement(geom, q("NoLine")).text = "0"
    if kind == "diamond":
        points = [(w / 2, 0), (w, h / 2), (w / 2, h), (0, h / 2), (w / 2, 0)]
    else:
        points = [(0, 0), (w, 0), (w, h), (0, h), (0, 0)]
        if kind in ("start", "output"):
            line = shape.find(q("Line"))
            SubElement(line, q("Rounding"), {"Unit": "IN_F"}).text = "0.16"
    first = SubElement(geom, q("MoveTo"), {"IX": "1"})
    SubElement(first, q("X"), {"Unit": "IN_F"}).text = str(points[0][0])
    SubElement(first, q("Y"), {"Unit": "IN_F"}).text = str(points[0][1])
    for idx, (px, py) in enumerate(points[1:], start=2):
        line_to = SubElement(geom, q("LineTo"), {"IX": str(idx)})
        SubElement(line_to, q("X"), {"Unit": "IN_F"}).text = str(px)
        SubElement(line_to, q("Y"), {"Unit": "IN_F"}).text = str(py)
    SubElement(shape, q("Text")).text = label
    return {"id": sid, "x": x, "y": y, "w": w, "h": h, "kind": kind, "label": label, "fill": fill, "stroke": stroke}


def add_connector(shapes, sid, points, label=""):
    min_x = min(x for x, _ in points)
    min_y = min(y for _, y in points)
    max_x = max(x for x, _ in points)
    max_y = max(y for _, y in points)
    w = max(max_x - min_x, 0.01)
    h = max(max_y - min_y, 0.01)
    shape = SubElement(shapes, q("Shape"), {
        "ID": str(sid), "NameU": f"Connector.{sid}", "Name": label or "连接线",
        "Type": "Shape", "LineStyle": "0", "FillStyle": "0", "TextStyle": "0"
    })
    add_xform(shape, min_x + w / 2, min_y + h / 2, w, h)
    line = SubElement(shape, q("Line"))
    SubElement(line, q("LineWeight"), {"Unit": "IN_F"}).text = "0.012"
    SubElement(line, q("LineColor")).text = "#667085"
    SubElement(line, q("LinePattern")).text = "1"
    SubElement(line, q("EndArrow")).text = "13"
    SubElement(line, q("EndArrowSize")).text = "2"
    fill = SubElement(shape, q("Fill"))
    SubElement(fill, q("FillPattern")).text = "0"
    add_text_style(shape, 0.105, True, "#475467")
    geom = SubElement(shape, q("Geom"), {"IX": "0"})
    SubElement(geom, q("NoFill")).text = "1"
    SubElement(geom, q("NoLine")).text = "0"
    first = SubElement(geom, q("MoveTo"), {"IX": "1"})
    SubElement(first, q("X"), {"Unit": "IN_F"}).text = str(points[0][0] - min_x)
    SubElement(first, q("Y"), {"Unit": "IN_F"}).text = str(points[0][1] - min_y)
    for idx, (px, py) in enumerate(points[1:], start=2):
        line_to = SubElement(geom, q("LineTo"), {"IX": str(idx)})
        SubElement(line_to, q("X"), {"Unit": "IN_F"}).text = str(px - min_x)
        SubElement(line_to, q("Y"), {"Unit": "IN_F"}).text = str(py - min_y)
    if label:
        SubElement(shape, q("Text")).text = label
    return {"points": points, "label": label}


def build_vdx():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    root = Element(q("VisioDocument"), {"xml:space": "preserve"})
    props = SubElement(root, q("DocumentProperties"))
    SubElement(props, q("Title")).text = "智能组表执行流程图（Word单页版）"
    SubElement(props, q("Subject")).text = "全局规则、个性化规则、字段映射、AI补充、报表组装及脚本复用"
    SubElement(props, q("Creator")).text = ""
    settings = SubElement(root, q("DocumentSettings"))
    SubElement(settings, q("GlueSettings")).text = "9"
    SubElement(settings, q("SnapSettings")).text = "31"
    faces = SubElement(root, q("FaceNames"))
    SubElement(faces, q("FaceName"), {"ID": "0", "Name": "Microsoft YaHei", "UnicodeRanges": "-1 -1 0 0", "CharSets": "134"})
    styles = SubElement(root, q("StyleSheets"))
    SubElement(styles, q("StyleSheet"), {"ID": "0", "NameU": "No Style", "Name": "No Style"})

    pages = SubElement(root, q("Pages"))
    page = SubElement(pages, q("Page"), {"ID": "0", "NameU": "Page-1", "Name": "智能组表执行流程"})
    page_sheet = SubElement(page, q("PageSheet"))
    page_props = SubElement(page_sheet, q("PageProps"))
    SubElement(page_props, q("PageWidth"), {"Unit": "IN_F"}).text = str(PAGE_W)
    SubElement(page_props, q("PageHeight"), {"Unit": "IN_F"}).text = str(PAGE_H)
    SubElement(page_props, q("PageScale"), {"Unit": "IN_F"}).text = "1"
    SubElement(page_props, q("DrawingScale"), {"Unit": "IN_F"}).text = "1"
    SubElement(page_props, q("DrawingSizeType")).text = "0"
    SubElement(page_props, q("DrawingScaleType")).text = "0"
    print_props = SubElement(page_sheet, q("PrintProps"))
    for tag in ("PageLeftMargin", "PageRightMargin", "PageTopMargin", "PageBottomMargin"):
        SubElement(print_props, q(tag), {"Unit": "IN_F"}).text = "0.35"
    SubElement(print_props, q("PrintPageOrientation")).text = "0"
    SubElement(print_props, q("PrintPageScale")).text = "100"

    shapes = SubElement(page, q("Shapes"))
    nodes = []
    nodes.append(add_node(shapes, 1, "Title", "智能组表执行流程图", 4.134, 11.3, 7.2, 0.38, fill="#FFFFFF", stroke="#FFFFFF"))
    nodes[-1]["title"] = True
    # Override title text size and weight.
    title_shape = shapes[-1]
    title_shape.find(q("Char")).find(q("Size")).text = "0.2222"
    title_shape.find(q("Char")).find(q("Style")).text = "1"

    nodes.extend([
        add_node(shapes, 2, "Start", "上传源文件和目标模板", 4.134, 10.55, kind="start", fill="#E8EEF5", stroke="#6B8FB3"),
        add_node(shapes, 3, "Rules", "加载全局规则和个性化规则", 4.134, 9.72),
        add_node(shapes, 4, "Parse", "解析文件结构和字段", 4.134, 8.88),
        add_node(shapes, 5, "Decision", "已有脚本可复用", 4.134, 7.98, 2.75, 0.70, "diamond", "#FFECBD", "#FFC943"),
        add_node(shapes, 6, "ArchiveRun", "执行已归档脚本", 1.62, 7.03, 2.25, 0.48, fill="#C6FAF6", stroke="#5AD8CC"),
        add_node(shapes, 7, "Query", "查询系统已有映射", 4.134, 7.03),
        add_node(shapes, 8, "Decision", "字段映射是否完整", 4.134, 6.08, 2.75, 0.70, "diamond", "#FFECBD", "#FFC943"),
        add_node(shapes, 9, "AI", "AI分析未匹配字段", 6.72, 5.18, 2.25, 0.48, fill="#DCCCFF", stroke="#874FFF"),
        add_node(shapes, 10, "MappingSave", "补充映射并写入映射表", 6.72, 4.35, 2.35, 0.48, fill="#C6FAF6", stroke="#5AD8CC"),
        add_node(shapes, 11, "Mapping", "形成完整映射关系", 4.134, 4.35),
        add_node(shapes, 12, "Assemble", "按映射关系组装报表", 4.134, 3.43),
        add_node(shapes, 13, "Script", "生成并归档Python脚本", 4.134, 2.50, fill="#C6FAF6", stroke="#5AD8CC"),
        add_node(shapes, 14, "Output", "生成终端格式报表", 4.134, 1.55, kind="output", fill="#CDF4D3", stroke="#66D575"),
    ])

    def top(n): return (n["x"], n["y"] + n["h"] / 2)
    def bottom(n): return (n["x"], n["y"] - n["h"] / 2)
    def left(n): return (n["x"] - n["w"] / 2, n["y"])
    def right(n): return (n["x"] + n["w"] / 2, n["y"])
    by_id = {n["id"]: n for n in nodes}
    connectors = []
    cid = 100
    for a, b in ((2, 3), (3, 4), (4, 5), (5, 7), (7, 8), (8, 11), (11, 12), (12, 13), (13, 14)):
        label = "否" if (a, b) == (5, 7) else ("是" if (a, b) == (8, 11) else "")
        connectors.append(add_connector(shapes, cid, [bottom(by_id[a]), top(by_id[b])], label))
        cid += 1
    # Script cache hit branch.
    start = left(by_id[5]); finish = top(by_id[6])
    connectors.append(add_connector(shapes, cid, [start, (by_id[6]["x"], start[1]), finish], "是")); cid += 1
    # Missing mappings branch and rejoin.
    start = right(by_id[8]); finish = top(by_id[9])
    connectors.append(add_connector(shapes, cid, [start, (by_id[9]["x"], start[1]), finish], "否")); cid += 1
    connectors.append(add_connector(shapes, cid, [bottom(by_id[9]), top(by_id[10])], "")); cid += 1
    connectors.append(add_connector(shapes, cid, [left(by_id[10]), right(by_id[11])], "")); cid += 1
    # Archived script bypasses mapping and assembly.
    start = bottom(by_id[6]); finish = left(by_id[14])
    connectors.append(add_connector(shapes, cid, [start, (0.52, start[1]), (0.52, finish[1]), finish], ""))

    ElementTree(root).write(OUT, encoding="utf-8", xml_declaration=True)
    return nodes, connectors


def build_preview(nodes, connectors):
    PREVIEW.parent.mkdir(parents=True, exist_ok=True)
    scale = 240
    img = Image.new("RGB", (int(PAGE_W * scale), int(PAGE_H * scale)), "white")
    draw = ImageDraw.Draw(img)
    font_path = Path(r"C:\Windows\Fonts\msyh.ttc")
    font = ImageFont.truetype(str(font_path), 16) if font_path.exists() else ImageFont.load_default()
    font_bold = ImageFont.truetype(str(font_path), 17) if font_path.exists() else font
    title_font = ImageFont.truetype(str(font_path), 26) if font_path.exists() else font_bold

    def xy(x, y): return int(x * scale), int((PAGE_H - y) * scale)
    for conn in connectors:
        pts = [xy(x, y) for x, y in conn["points"]]
        draw.line(pts, fill="#667085", width=2, joint="curve")
        if len(pts) >= 2:
            x2, y2 = pts[-1]; x1, y1 = pts[-2]
            if abs(x2 - x1) > abs(y2 - y1):
                sign = 1 if x2 > x1 else -1
                arrow = [(x2, y2), (x2 - sign * 10, y2 - 5), (x2 - sign * 10, y2 + 5)]
            else:
                sign = 1 if y2 > y1 else -1
                arrow = [(x2, y2), (x2 - 5, y2 - sign * 10), (x2 + 5, y2 - sign * 10)]
            draw.polygon(arrow, fill="#667085")
        if conn["label"]:
            mid = pts[len(pts) // 2]
            draw.text((mid[0] + 5, mid[1] - 18), conn["label"], fill="#475467", font=font_bold)
    for node in nodes:
        x, y = xy(node["x"], node["y"])
        w, h = int(node["w"] * scale), int(node["h"] * scale)
        box = (x - w // 2, y - h // 2, x + w // 2, y + h // 2)
        if node.get("title"):
            draw.text((x, y), node["label"], anchor="mm", fill="#24292F", font=title_font)
            continue
        if node["kind"] == "diamond":
            draw.polygon([(x, box[1]), (box[2], y), (x, box[3]), (box[0], y)], fill=node["fill"], outline=node["stroke"], width=2)
        else:
            radius = 14 if node["kind"] in ("start", "output") else 2
            draw.rounded_rectangle(box, radius=radius, fill=node["fill"], outline=node["stroke"], width=2)
        draw.text((x, y), node["label"], anchor="mm", fill="#24292F", font=font_bold if node["kind"] in ("start", "output") else font)
    img.save(PREVIEW)


def build_pdf():
    PDF_OUT.parent.mkdir(parents=True, exist_ok=True)
    page_w, page_h = A4
    pdf = canvas.Canvas(str(PDF_OUT), pagesize=A4)
    pdf.setTitle("智能组表执行流程图（单页版）")
    pdf.drawImage(str(PREVIEW), 0, 0, width=page_w, height=page_h, preserveAspectRatio=True, mask="auto")
    pdf.showPage()
    pdf.save()


if __name__ == "__main__":
    ns, cs = build_vdx()
    build_preview(ns, cs)
    build_pdf()
    print(OUT)
    print(PREVIEW)
    print(PDF_OUT)
