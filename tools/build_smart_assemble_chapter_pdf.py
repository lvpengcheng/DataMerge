from __future__ import annotations

import io
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    Flowable,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
SOURCE = Path(r"C:\Users\Administrator\Desktop\服务交付报表效能优化 概要设计 731 (1).pdf")
OUT_DIR = ROOT / "output" / "pdf"
CHAPTER_PDF = ROOT / "tmp" / "pdfs" / "smart_assemble_chapter.pdf"
OVERLAY_PDF = ROOT / "tmp" / "pdfs" / "page27_overlay.pdf"
OUTPUT = OUT_DIR / "服务交付报表效能优化_概要设计_补充智能组表.pdf"


FONT_REGULAR = "MicrosoftYaHei"
FONT_BOLD = "MicrosoftYaHei-Bold"
pdfmetrics.registerFont(TTFont(FONT_REGULAR, r"C:\Windows\Fonts\msyh.ttc", subfontIndex=0))
pdfmetrics.registerFont(TTFont(FONT_BOLD, r"C:\Windows\Fonts\msyhbd.ttc", subfontIndex=0))


INK = colors.HexColor("#20252B")
MUTED = colors.HexColor("#69717D")
BLUE = colors.HexColor("#0B63F6")
LIGHT_BLUE = colors.HexColor("#EEF5FF")
GRID = colors.HexColor("#D5DAE0")
HEADER_BG = colors.HexColor("#F0F2F4")
GREEN_BG = colors.HexColor("#ECF8F0")
ORANGE_BG = colors.HexColor("#FFF5E8")


def styles():
    s = getSampleStyleSheet()
    return {
        "h1": ParagraphStyle(
            "h1", parent=s["Heading1"], fontName=FONT_BOLD, fontSize=20,
            leading=28, textColor=INK, spaceAfter=14,
        ),
        "h2": ParagraphStyle(
            "h2", parent=s["Heading2"], fontName=FONT_BOLD, fontSize=15,
            leading=22, textColor=INK, spaceBefore=5, spaceAfter=9,
        ),
        "h3": ParagraphStyle(
            "h3", parent=s["Heading3"], fontName=FONT_BOLD, fontSize=12,
            leading=18, textColor=INK, spaceBefore=4, spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "body", parent=s["BodyText"], fontName=FONT_REGULAR, fontSize=9.6,
            leading=16, textColor=INK, alignment=TA_LEFT, spaceAfter=6,
        ),
        "small": ParagraphStyle(
            "small", parent=s["BodyText"], fontName=FONT_REGULAR, fontSize=8.3,
            leading=13, textColor=INK,
        ),
        "tiny": ParagraphStyle(
            "tiny", parent=s["BodyText"], fontName=FONT_REGULAR, fontSize=7.4,
            leading=11.2, textColor=INK,
        ),
        "callout": ParagraphStyle(
            "callout", parent=s["BodyText"], fontName=FONT_REGULAR, fontSize=9.2,
            leading=15, textColor=MUTED, leftIndent=9, rightIndent=7,
        ),
        "center": ParagraphStyle(
            "center", parent=s["BodyText"], fontName=FONT_BOLD, fontSize=8.4,
            leading=12, textColor=INK, alignment=TA_CENTER,
        ),
    }


S = styles()


def P(text: str, style: str = "body") -> Paragraph:
    return Paragraph(text, S[style])


def bullet(text: str) -> Paragraph:
    return P(f'<font color="#0B63F6">•</font>&nbsp;&nbsp;{text}')


def callout(title: str, text: str, bg=LIGHT_BLUE) -> Table:
    box = Table([[P(f"<b>{title}</b><br/>{text}", "callout")]], colWidths=[174 * mm])
    box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), bg),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#C7D8F2")),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return box


def grid_table(rows, widths, header=True, font="small") -> Table:
    data = []
    for r_idx, row in enumerate(rows):
        style = "center" if header and r_idx == 0 else font
        data.append([P(str(v), style) for v in row])
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0, hAlign="LEFT")
    commands = [
        ("GRID", (0, 0), (-1, -1), 0.45, GRID),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]
    if header:
        commands.extend([
            ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
            ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.HexColor("#C7CDD4")),
        ])
    t.setStyle(TableStyle(commands))
    return t


class FlowRow(Flowable):
    def __init__(self, labels, width=174 * mm, height=25 * mm):
        super().__init__()
        self.labels = labels
        self.width = width
        self.height = height

    def wrap(self, avail_width, avail_height):
        return min(self.width, avail_width), self.height

    def draw(self):
        c = self.canv
        n = len(self.labels)
        gap = 6 * mm
        total_gap = gap * (n - 1)
        box_w = (self.width - total_gap) / n
        box_h = 15 * mm
        y = 5 * mm
        for i, label in enumerate(self.labels):
            x = i * (box_w + gap)
            c.setFillColor(LIGHT_BLUE if i not in (0, n - 1) else colors.HexColor("#F3F7FC"))
            c.setStrokeColor(colors.HexColor("#AAC5EE"))
            c.roundRect(x, y, box_w, box_h, 3, fill=1, stroke=1)
            c.setFillColor(INK)
            c.setFont(FONT_BOLD, 8)
            lines = label.split("\n")
            for j, line in enumerate(lines):
                c.drawCentredString(x + box_w / 2, y + box_h / 2 + 2 - j * 10, line)
            if i < n - 1:
                ax = x + box_w + 1.2 * mm
                ay = y + box_h / 2
                c.setStrokeColor(BLUE)
                c.setFillColor(BLUE)
                c.line(ax, ay, ax + gap - 2.4 * mm, ay)
                c.line(ax + gap - 2.4 * mm, ay, ax + gap - 4 * mm, ay + 1.5 * mm)
                c.line(ax + gap - 2.4 * mm, ay, ax + gap - 4 * mm, ay - 1.5 * mm)


def page_decor(c: canvas.Canvas, doc):
    c.saveState()
    c.setStrokeColor(colors.HexColor("#E8EBEF"))
    c.line(22 * mm, 15 * mm, 188 * mm, 15 * mm)
    c.setFont(FONT_REGULAR, 7.5)
    c.setFillColor(MUTED)
    c.drawString(22 * mm, 10 * mm, "服务交付报表效能优化 - 智能组表补充设计")
    c.drawRightString(188 * mm, 10 * mm, str(doc.page))
    c.restoreState()


def build_chapter():
    CHAPTER_PDF.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(CHAPTER_PDF), pagesize=A4,
        leftMargin=22 * mm, rightMargin=22 * mm,
        topMargin=18 * mm, bottomMargin=21 * mm,
        title="智能组表概要设计补充章节",
        author="Codex",
    )
    story = []

    story += [P("6.1 系统功能说明", "h1")]
    story += [P(
        "智能组表用于解决<b>数据源不固定、输出报表结构固定</b>的报表生产场景。用户每次可上传数量、文件名、sheet 名和字段命名均可能变化的 Excel 数据源，并指定一个固定输出模板。系统依据组表规则识别各数据表角色、人员主体和字段语义，将数据写入模板规定区域，最终输出结构、列顺序和样式稳定的 Excel 报表。"
    )]
    story += [grid_table([
        ["功能点", "功能描述", "产出"],
        ["规则与模板约束", "选择全局或客户级组表规则；以用户上传模板的激活 sheet 作为固定输出结构。", "本次任务规则上下文、模板结构"],
        ["非固定源解析", "支持多文件及全部可见 sheet；识别表头、列字母、公式列、格式及少量脱敏样例。", "标准化源结构描述"],
        ["主体与字段匹配", "先确定人员集合及关联主键，再结合文件名、sheet 名、列名和规则完成字段映射。", "结构化字段映射清单"],
        ["固定模板组装", "保留模板既有结构与样式，按需扩展数据行，将多源数据关联、合并后写入固定列。", "固定格式原版报表"],
        ["结果复核与复用", "记录任务映射；正确结果提升映射置信度，错误结果允许逐列修正并重新组表。", "纯值版、知识库、代码存档"],
    ], [32*mm, 96*mm, 46*mm])]
    story += [Spacer(1, 4 * mm), P("适用范围", "h2")]
    story += [bullet("源文件可以按月份、地区或客户变化，允许多文件、多 sheet 和不同字段命名。")]
    story += [bullet("输出模板在一次任务内固定；输出列、顺序、表头、样式和公式要求以模板为准。")]
    story += [bullet("优先覆盖人事、薪资、社保、公积金等按人员主键汇总的二维明细表。")]
    story += [Spacer(1, 2 * mm), callout(
        "设计边界",
        "本章不描述智算平台的通用训练、脚本计算等能力，仅描述智能组表如何将不定数据源转换为固定输出报表。对于无法识别主键、源表存在严重结构损坏或规则本身冲突的情况，系统应停止自动填充并提示人工处理。",
        ORANGE_BG,
    )]

    story += [PageBreak(), P("6.2 逻辑模型", "h1")]
    story += [P("智能组表由规则配置、任务执行、字段映射知识三个核心对象组成，文件实体按租户隔离存储。")]
    story += [FlowRow(["规则文件\n固定模板", "结构解析\n签名计算", "字段映射\n组装执行", "原版/纯值版\n结果反馈"]), Spacer(1, 4*mm)]
    story += [P("1. 业务对象", "h2")]
    story += [grid_table([
        ["对象", "关键属性", "说明"],
        ["组表规则 assemble_rules", "名称、作用域、租户、附件、说明、上传人", "全局规则供所有租户选择；客户规则仅指定租户可见。"],
        ["组表任务 assemble_tasks", "租户、规则、结构签名、状态、代码路径、映射、输出文件、反馈", "记录一次从上传、分析、执行到复核的完整链路。"],
        ["字段映射 assemble_field_mappings", "租户、源列、目标列、模板签名、匹配类型、确认次数、状态", "沉淀同类模板下可复用的字段对应关系。"],
    ], [42*mm, 72*mm, 60*mm])]
    story += [Spacer(1, 4*mm), P("2. 关键关系", "h2")]
    story += [bullet("一个规则可被多个组表任务引用；规则删除时历史任务保留，关联置空。")]
    story += [bullet("一个任务对应一个租户、一个固定模板和一个或多个源文件，输出一个原版及可选纯值版。")]
    story += [bullet("字段映射以 tenant_id + template_signature + source_column 唯一，避免跨客户、跨模板误复用。")]
    story += [bullet("任务保存完整 field_mapping 和 corrected_mapping，用于问题追溯及人工修正重跑。")]
    story += [Spacer(1, 3*mm), P("3. 文件存储", "h2")]
    story += [grid_table([
        ["内容", "存储位置", "隔离策略"],
        ["全局组表规则", "global_assets/assemble_rules/", "平台级只读复用"],
        ["客户组表规则", "tenants/{tenant_id}/assemble_rules/", "按租户隔离"],
        ["代码存档", "tenants/{tenant_id}/assemble_scripts/{signature}.py", "结构和规则一致才命中"],
        ["任务结果及归档输入", "tenants/{tenant_id}/assemble_results/{task_id}/", "任务级隔离并支持复核重跑"],
    ], [38*mm, 91*mm, 45*mm])]

    story += [PageBreak(), P("6.3 系统功能权限", "h1")]
    story += [grid_table([
        ["功能 / 角色", "普通业务用户", "规则管理员", "平台管理员"],
        ["查看可用规则", "√ 本租户 + 全局", "√", "√"],
        ["提交组表任务、查看进度、下载结果", "√ 本租户", "√ 本租户", "√ 授权租户"],
        ["查看历史、结果反馈、修正映射重跑", "√ 本人/本租户", "√ 本租户", "√ 授权租户"],
        ["上传、替换、删除组表规则", "-", "√", "√"],
        ["查看、停用、恢复、删除字段知识", "-", "√", "√"],
    ], [68*mm, 35*mm, 35*mm, 36*mm])]
    story += [Spacer(1, 3*mm), callout(
        "权限标识",
        "任务执行使用 tools.assemble；规则及知识库管理使用 tools.assemble.manage。所有任务、存档代码、映射知识和输出文件均需校验用户可操作租户范围。",
    )]
    story += [Spacer(1, 5*mm), P("6.4 接口定义", "h1")]
    story += [grid_table([
        ["接口", "方法", "用途"],
        ["/api/assemble/rules", "GET / POST", "查询可用规则；管理员上传全局或客户规则。"],
        ["/api/assemble/rules/{id}/upload", "POST", "替换指定规则附件。"],
        ["/api/assemble/submit", "POST", "提交源文件、固定模板、规则、AI 提供者及强制重匹配标识。"],
        ["/api/assemble/tasks/{id}/stream", "GET", "SSE 返回解析、匹配、生成、执行和完成事件。"],
        ["/api/assemble/tasks/{id}/status", "GET", "查询任务状态、输出文件及错误信息。"],
        ["/api/assemble/download/{id}/{file}", "GET", "下载任务原版或纯值版结果。"],
        ["/api/assemble/history", "GET", "按当前用户授权租户查询任务历史。"],
        ["/api/assemble/tasks/{id}/feedback", "POST", "确认正确或标记有误，驱动映射置信度状态。"],
        ["/api/assemble/tasks/{id}/rematch", "POST", "使用人工修正映射和归档输入重新组表。"],
        ["/api/assemble/mappings", "GET / PUT / DELETE", "管理员维护字段映射知识库。"],
    ], [82*mm, 25*mm, 67*mm], font="tiny")]

    story += [PageBreak(), P('<font name="Helvetica-Bold">6.5</font> 数据初始化', "h1")]
    story += [bullet("初始化权限点 tools.assemble 和 tools.assemble.manage，并分配至对应业务角色。")]
    story += [bullet("初始化智能组表全局规则，至少包含组表模式、人员主体、关联主键、字段冲突和空值策略。")]
    story += [bullet("创建 assemble_rules、assemble_tasks、assemble_field_mappings 三类表及唯一约束、索引。")]
    story += [bullet("建立全局规则、租户规则、代码存档和任务结果目录；设置容量、留存和清理策略。")]
    story += [Spacer(1, 5*mm), P("6.6 重点功能设计", "h1")]
    story += [P("6.6.1 整体业务流程", "h2")]
    story += [FlowRow(["上传源文件\n选择固定模板", "解析结构\n计算签名", "复用或生成\n字段映射", "沙箱组装\n输出复核"]), Spacer(1, 3*mm)]
    process_rows = [
        ["步骤", "处理说明", "控制点"],
        ["1. 提交", "选择规则，上传多个非固定源文件和一个固定输出模板。", "文件类型、租户权限、加密检测"],
        ["2. 解析", "读取全部可见源 sheet；仅以模板激活 sheet 作为目标区域。", "样例数据脱敏；不改变真实执行数据"],
        ["3. 签名", "按源列集合、模板表头、规则内容生成稳定签名。", "文件名或月份变化不应导致无效失配"],
        ["4. 匹配", "优先命中代码存档，其次采用已确认映射，再由 AI 处理剩余字段。", "歧义不猜测；未确认映射不自动采用"],
        ["5. 组装", "先对齐人员主体，再按主键关联、纵向拼接或两阶段混合组装。", "主键格式归一、去重、冲突优先级"],
        ["6. 输出", "复制模板、填充数据、按需扩展样式行，生成原版和纯值版。", "固定列顺序、表头、样式和公式"],
        ["7. 复核", "用户确认正确或逐列修正映射并重跑。", "结果闭环更新知识库和代码存档"],
    ]
    story += [grid_table(process_rows, [24*mm, 99*mm, 51*mm], font="tiny")]

    story += [PageBreak(), P("6.6.2 非固定数据源解析与固定模板识别", "h1")]
    story += [P("源文件与模板采用不同解析策略，以确保输入可以变化而输出边界稳定。")]
    story += [grid_table([
        ["项目", "非固定数据源", "固定输出模板"],
        ["读取范围", "多个 Excel 文件、全部可见 sheet", "单个模板文件、激活 sheet"],
        ["提取内容", "文件名、sheet 名、表头、列字母、公式、格式、前 3 行脱敏样例", "表头行、数据区域、列名、列字母、公式列和列格式"],
        ["变化容忍", "允许文件名、sheet 名、列顺序和同义字段变化", "表头结构决定输出签名；sheet 按月改名不影响识别"],
        ["异常策略", "无表头时使用解析兜底并记录日志", "只有表头无数据行时停止并明确提示"],
    ], [30*mm, 72*mm, 72*mm])]
    story += [Spacer(1, 4*mm), P("结构签名", "h2")]
    story += [bullet("源签名：对全部源文件的列名集合进行稳定排序后计算摘要，忽略无业务意义的文件排列变化。")]
    story += [bullet("模板签名：基于激活 sheet 的有效表头和关键列集合计算，不包含 sheet 名。")]
    story += [bullet("总签名：规则内容哈希 + 源签名 + 模板签名。命中且未强制重匹配时，可直接执行已验证代码。")]
    story += [Spacer(1, 3*mm), P("6.6.3 组表模式与字段匹配", "h1")]
    story += [grid_table([
        ["组表模式", "识别特征", "处理方式"],
        ["同构整合", "多个源表列结构高度相似，如不同地区同类明细。", "纵向拼接；同主键数值加总、文本取首值并去重。"],
        ["职责分工", "各源表字段差异大，分别提供花名册、考勤、社保等信息。", "确定人员主体后，按主键关联各表负责字段。"],
        ["混合模式", "部分表同构，另有入职、离职或专项补充表。", "先同构合并形成中间结果，再与职责表关联。"],
    ], [32*mm, 68*mm, 74*mm])]
    story += [Spacer(1, 3*mm), callout(
        "匹配原则",
        "先确定人员主体和关联主键，再匹配字段。主键优先级为身份证/证件号 > 工号/员工编号 > 姓名+部门 > 姓名。无法确定主键或存在多人重名时宁可留空并告警，不得猜测填充。",
        GREEN_BG,
    )]

    story += [PageBreak(), P('<font name="Helvetica-Bold">6.6.4</font> 映射复用与置信度闭环', "h1")]
    story += [grid_table([
        ["层级", "命中条件", "处理"],
        ["代码存档", "规则、源结构和模板结构总签名一致", "跳过 AI，直接执行已验证代码；强制重匹配时跳过。"],
        ["已确认语义映射", "同租户、同模板签名、源列一致，状态 active 且确认次数达到阈值", "自动采用并记录 used_mapping_ids。"],
        ["同名列", "源列名与模板有效列名完全一致", "确定性白名单直接采用，不受错误反馈连坐。"],
        ["AI 候选映射", "未被前述层级覆盖的字段", "生成结构化映射；成功执行后以 pending 状态入库。"],
        ["冲突或否定映射", "同源列指向多个目标，或用户标记错误", "置为 review_needed，后续不自动采用。"],
    ], [34*mm, 76*mm, 64*mm])]
    story += [Spacer(1, 4*mm), bullet("语义映射需累计 2 次人工确认后转为 active，避免一次错误结果污染后续任务。")]
    story += [bullet("用户可在结果页逐列修正映射；修正后复用归档源文件和模板重跑，并将修正映射持久化到代码存档。")]
    story += [Spacer(1, 5*mm), P("6.6.5 执行与固定输出", "h1")]
    story += [bullet("执行代码在独立子进程和受控沙箱中运行，避免 Excel 计算阻塞 API 事件循环。")]
    story += [bullet("复制固定模板后填充目标数据行；当数据行超出模板预留范围时复制最后数据行样式向下扩展。")]
    story += [bullet("原版保留源数据 sheet 和新生成公式，便于追溯；纯值版将新增公式计算为值并移除源 sheet，便于交付。")]
    story += [bullet("结果按租户和任务落盘，下载时校验任务归属和文件名，禁止跨租户访问。")]
    story += [Spacer(1, 4*mm), P("6.6.6 异常、安全与验收", "h1")]
    story += [grid_table([
        ["类别", "设计要求", "验收要点"],
        ["数据安全", "仅脱敏样例进入 AI 上下文；真实数据仅在本地沙箱执行。", "姓名、证件、手机、邮箱、金额等日志样例不可明文泄露。"],
        ["执行异常", "AI 生成失败自动带错误重试 1 次；执行失败记录 error 并支持重试。", "失败代码和未经验证映射不得写入存档或知识库。"],
        ["过程可见", "SSE 展示解析、匹配、生成、执行和完成状态。", "长任务期间日志与心跳持续可见，断线后可查询最终状态。"],
        ["固定输出", "输出列、顺序、表头、样式与模板一致，人员及金额按规则组装。", "典型同构、职责分工、混合场景均通过业务样例核对。"],
    ], [29*mm, 80*mm, 65*mm], font="tiny")]

    doc.build(story, onFirstPage=page_decor, onLaterPages=page_decor)


def build_page27_overlay(width: float, height: float):
    OVERLAY_PDF.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OVERLAY_PDF), pagesize=(width, height))
    # 清理原文第 6 章空占位及紧随其后的附录标题，保留上方监控告警表。
    c.setFillColor(colors.white)
    c.rect(0, 0, width, 325, fill=1, stroke=0)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 15)
    c.drawString(44, 300, "6")
    c.setFont(FONT_BOLD, 15)
    c.drawString(58, 300, "智能组表")
    c.setFont(FONT_REGULAR, 9.6)
    text = c.beginText(44, 277)
    text.setLeading(16)
    for line in [
        "智能组表用于实现“数据源不固定、输出报表格式固定”的自动化组装。系统读取多个结构可变的",
        "Excel 数据源，以用户指定的固定模板为输出边界，通过规则约束、结构识别、主体对齐和字段",
        "映射完成报表填充，并以人工复核结果持续提升同类任务的复用准确率。",
    ]:
        text.textLine(line)
    c.drawText(text)
    c.setFillColor(MUTED)
    c.setFont(FONT_REGULAR, 8.5)
    c.drawString(44, 217, "本章仅描述智能组表能力，不扩展智算平台其他训练与计算模块。")
    c.save()


def merge_pdf():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    source = PdfReader(str(SOURCE))
    width = float(source.pages[26].mediabox.width)
    height = float(source.pages[26].mediabox.height)
    build_page27_overlay(width, height)

    overlay = PdfReader(str(OVERLAY_PDF)).pages[0]
    chapter = PdfReader(str(CHAPTER_PDF))
    writer = PdfWriter()
    for idx in range(26):
        writer.add_page(source.pages[idx])

    page27 = source.pages[26]
    page27.merge_page(overlay)
    writer.add_page(page27)

    for page in chapter.pages:
        writer.add_page(page)

    for idx in range(27, len(source.pages)):
        writer.add_page(source.pages[idx])

    writer.add_metadata({
        "/Title": "服务交付报表效能优化 概要设计 - 补充智能组表",
        "/Subject": "智能组表：非固定数据源到固定输出报表",
    })
    with OUTPUT.open("wb") as f:
        writer.write(f)


if __name__ == "__main__":
    build_chapter()
    merge_pdf()
    print(OUTPUT)
