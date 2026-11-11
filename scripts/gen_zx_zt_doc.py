# -*- coding: utf-8 -*-
"""单位-ZX主楼综合维修改造项目 办公用房维修项目初步设计（代项目建议书、可行性研究报告）生成脚本
格式：仿宋_GB2312、三节页码规则、目录TOC域（Heading 1 + TOC1）、第一章/第八章/第九章表格化。
封面按 2026-09-03 新标准：无眉题行、主标题26pt、副题"初步设计"26pt、"（代...）"22pt、底部16pt信息行含设计单位留白、编制时间固定"二〇二六年八月"。"""
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_LINE_SPACING, WD_BREAK
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

PATH = r'<工作目录>\单位-ZX主楼综合维修改造项目办公用房维修项目初步设计（代项目建议书、可行性研究报告）.docx'
FONT = '仿宋_GB2312'

doc = Document()

# ============ 页面设置 ============
sec0 = doc.sections[0]
sec0.page_width = Cm(21.0)
sec0.page_height = Cm(29.7)
sec0.top_margin = Cm(3.0)
sec0.bottom_margin = Cm(2.5)
sec0.left_margin = Cm(2.8)
sec0.right_margin = Cm(2.6)

normal = doc.styles['Normal']
normal.font.name = FONT
normal.font.size = Pt(12)
normal.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
normal.paragraph_format.line_spacing = Pt(24)
normal.paragraph_format.space_before = Pt(0)
normal.paragraph_format.space_after = Pt(0)

h1s = doc.styles['Heading 1']
h1s.font.name = FONT
h1s.font.size = Pt(16)
h1s.font.bold = True
h1s.font.color.rgb = RGBColor(0, 0, 0)
h1s.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
h1s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
h1s.paragraph_format.space_before = Pt(12)
h1s.paragraph_format.space_after = Pt(8)
h1s.paragraph_format.line_spacing = Pt(26)


def set_run_font(run, size=None, bold=None, font=FONT):
    run.font.name = font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)
    rFonts.set(qn('w:eastAsia'), font)
    rFonts.set(qn('w:cs'), font)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    return run


def para(text='', size=12, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True,
         line=24, before=0, after=0):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    if indent:
        pf.first_line_indent = Pt(size * 2)
    pf.line_spacing = Pt(line)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if text:
        set_run_font(p.add_run(text), size=size, bold=bold)
    return p


def para_design_unit(text, size=16, bold=True, line=34):
    """封面"设计单位"行专用：两端对齐(both) + 首行缩进 4 字符（firstLineChars=400）。
    区别于其余信息行的居中(CENTER)，2026-09-03 用户约定。"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = Pt(line)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    set_run_font(p.add_run(text), size=size, bold=bold)
    # 精确设置首行缩进 4 字符（w:firstLineChars=400）
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:firstLineChars'), '400')
    ind.set(qn('w:firstLine'), '1285')
    return p


def heading(text, size=14, before=6, after=6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.line_spacing = Pt(26)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.keep_with_next = True
    set_run_font(p.add_run(text), size=size, bold=True)
    return p


def h1(text):
    p = doc.add_paragraph(style='Heading 1')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.line_spacing = Pt(26)
    pf.space_before = Pt(12)
    pf.space_after = Pt(8)
    pf.keep_with_next = True
    set_run_font(p.add_run(text), size=16, bold=True)
    return p


def h2(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.line_spacing = Pt(24)
    pf.space_before = Pt(6)
    pf.space_after = Pt(4)
    pf.keep_with_next = True
    set_run_font(p.add_run(text), size=14, bold=True)
    return p


def h3(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.line_spacing = Pt(24)
    pf.space_before = Pt(4)
    pf.space_after = Pt(2)
    pf.keep_with_next = True
    set_run_font(p.add_run(text), size=12, bold=True)
    return p


def add_table(headers, rows, widths=None, font_size=10.5):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, htext in enumerate(headers):
        cell = t.rows[0].cells[j]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_font(cell.paragraphs[0].add_run(htext), size=font_size, bold=True)
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:fill'), 'E7E6E6')
        tcPr.append(shd)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.rows[i + 1].cells[j]
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run_font(cell.paragraphs[0].add_run(str(val)), size=font_size)
    if widths:
        for j, w in enumerate(widths):
            for row in t.rows:
                row.cells[j].width = Cm(w)
    return t


def add_field(paragraph, instr, cache='1'):
    r = paragraph.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = instr
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = cache
    f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
    for el in (f1, it, f2, t, f3):
        r._r.append(el)
    set_run_font(r)
    return r


# ===================== 封面（节1）—— 2026-09-03 新标准 =====================
for _ in range(4):
    para('', indent=False)
para('单位-ZX主楼综合维修改造项目', size=26, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=40)
para('', indent=False)
para('初步设计', size=26, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=40)
para('（代项目建议书、可行性研究报告）', size=22, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=36)
for _ in range(11):
    para('', indent=False)
para('项目单位：单位-ZX', size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=34)
para('建设地点：四川省地址-ZX', size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=34)
para_design_unit('设计单位：' + '\u3000' * 16)
para('编制时间：二〇二六年八月', size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=34)

sec1 = doc.add_section(WD_SECTION.NEW_PAGE)
sec1.page_width = Cm(21.0); sec1.page_height = Cm(29.7)
sec1.top_margin = Cm(3.0); sec1.bottom_margin = Cm(2.5)
sec1.left_margin = Cm(2.8); sec1.right_margin = Cm(2.6)

# ===================== 编制说明（节2） =====================
heading('编 制 说 明', size=22, before=6, after=12)
bs_items = [
    '一、本初步设计（代项目建议书、可行性研究报告）依据《四川省政府投资管理办法》（川府规〔2022〕5号）规定，按照"三合一"评审（项目建议书、可行性研究报告、初步设计）要求编制，用于单位-ZX主楼综合维修改造项目办公用房维修项目的审批。',
    '二、项目单位：单位-ZX。项目建设地点位于四川省地址-ZX。',
    '三、本文件主要内容包括：项目基本概况，项目现状情况及维修改造必要性，项目实施可行性论证，维修设计说明，安全设施设计、节能设计，环境保护及影响评价，项目实施组织管理，投资概算与资金筹措等。',
    '四、编制依据：《四川省政府投资管理办法》（川府规〔2022〕5号）、《政府投资项目可行性研究报告编写通用大纲》（2023年版）、《建筑工程设计文件编制深度规定》（2016年版）、《四川省省级机关办公用房维修管理实施办法（2023年修订版）》（川机管发〔2023〕163号）、《党政机关办公用房建设标准》（发改投资〔2014〕2674号）以及国家、四川省现行有关工程建设标准、规范及规定。',
    '五、项目投资概算依据设计图纸、现行工程量清单计价规范和四川省建设工程工程量清单计价定额等编制，项目概算总投资154.04万元。',
    '六、本初步设计说明（代项目建议书、可行性研究报告）、初步设计文件、投资概算书及申请审批函的内容保持一致。',
]
for it in bs_items:
    para(it, size=12)
for _ in range(2):
    para('', indent=False)
para('二〇二六年八月', size=14, align=WD_ALIGN_PARAGRAPH.RIGHT, indent=False, before=18)

sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
sec2.page_width = Cm(21.0); sec2.page_height = Cm(29.7)
sec2.top_margin = Cm(3.0); sec2.bottom_margin = Cm(2.5)
sec2.left_margin = Cm(2.8); sec2.right_margin = Cm(2.6)

# ===================== 目录（节3） =====================
heading('目    录', size=22, before=6, after=12)

TOC_ENTRIES = [
    ('第一章  项目基本概况', 1),
    ('第二章  项目现状情况及维修改造必要性', 2),
    ('第三章  项目实施可行性论证', 3),
    ('第四章  维修设计说明', 4),
    ('第五章  安全设施设计', 5),
    ('第六章  节能设计', 6),
    ('第七章  环境保护及影响评价', 7),
    ('第八章  项目实施组织管理', 8),
    ('第九章  投资概算与资金筹措', 9),
    ('第十章  结论与建议', 10),
]
EST_PAGES = {'第一章  项目基本概况': 2, '第二章  项目现状情况及维修改造必要性': 3,
             '第三章  项目实施可行性论证': 4, '第四章  维修设计说明': 5, '第五章  安全设施设计': 8,
             '第六章  节能设计': 9, '第七章  环境保护及影响评价': 9, '第八章  项目实施组织管理': 10,
             '第九章  投资概算与资金筹措': 11, '第十章  结论与建议': 13}

p_toc_begin = doc.add_paragraph()
p_toc_begin.alignment = WD_ALIGN_PARAGRAPH.LEFT
r = p_toc_begin.add_run()
f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin'); f1.set(qn('w:dirty'), 'true')
it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = r' TOC \o "1-1" \h \z \u '
f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
t = OxmlElement('w:t'); t.text = '目录'
for el in (f1, it, f2, t):
    r._r.append(el)
set_run_font(r, size=12)

toc_paras = []
for title, chap in TOC_ENTRIES:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = Pt(24)
    pf.tab_stops.add_tab_stop(Cm(15.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    anchor = f'_TOC_CS_{chap}'
    hl = OxmlElement('w:hyperlink')
    hl.set(qn('w:anchor'), anchor)
    r1 = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    rf.set(qn('w:ascii'), FONT); rf.set(qn('w:eastAsia'), FONT); rf.set(qn('w:hAnsi'), FONT); rf.set(qn('w:cs'), FONT)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '24')
    rPr.append(rf); rPr.append(sz)
    t1 = OxmlElement('w:t'); t1.text = title; t1.set(qn('xml:space'), 'preserve')
    r1.append(rPr); r1.append(t1)
    hl.append(r1)
    p._p.append(hl)
    tab_r = p.add_run('\t')
    set_run_font(tab_r, size=12)
    pg_r = p.add_run(str(EST_PAGES.get(title, '')))
    set_run_font(pg_r, size=12)
    toc_paras.append(p)

p_toc_end = toc_paras[-1]
r = p_toc_end.add_run()
f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
r._r.append(f3)
set_run_font(r, size=12)

brk = doc.add_paragraph()
brk.add_run().add_break(WD_BREAK.PAGE)


def add_bookmark(paragraph, name, bid):
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), str(bid))
    start.set(qn('w:name'), name)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str(bid))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


# ===================== 正文 =====================

# 第一章
p = h1('第一章  项目基本概况')
add_bookmark(p, '_TOC_CS_1', 101)
h2('1.1  项目基本情况')
add_table(
    ['项目', '内容'],
    [
        ['项目名称', '单位-ZX主楼综合维修改造项目（办公用房维修项目）'],
        ['项目单位', '单位-ZX'],
        ['建设地点', '四川省地址-ZX'],
        ['项目性质', '党政机关办公用房维修项目'],
        ['维修内容', '主楼房门/防火门更换、地下室车库入口挡土墙漏水整改翻新、议政厅雨棚维修'],
        ['建设工期', '180日历天'],
        ['投资概算', '项目概算总投资154.04万元'],
        ['资金来源', '申请省级财政资金'],
    ],
    widths=[3.0, 12.4],
)
para('', indent=False)
h2('1.2  建筑基本情况')
add_table(
    ['项目', '内容'],
    [
        ['建筑名称', '单位-ZX主楼'],
        ['结构形式', '钢筋混凝土框架结构'],
        ['建成年代', '2003年度建成并投入使用'],
        ['已使用年限', '截至2026年已使用23年'],
        ['使用性质', '办公用房'],
        ['坐落位置', '地址-ZX'],
        ['维修范围', '主楼办公用房及会议室、公共区域门，议政厅屋面雨棚，地下室车库入口挡土墙等部位'],
        ['维修性质', '不改变既有建筑结构体系和使用性质，仅进行既有办公用房及附属设施维修改造'],
    ],
    widths=[3.0, 12.4],
)
para('', indent=False)
h2('1.3  项目维修内容及规模')
para('本项目维修内容及规模如下：')
add_table(
    ['序号', '分项工程', '主要工程内容及规模'],
    [
        ['1', '主楼房门/防火门更换', '电子密码锁130个；木质门带套（单开466.83㎡、双开81.32㎡）；防火门（单开乙级65.52㎡、双开乙级88.2㎡、单开甲级1.89㎡）；防盗门（单开13.23㎡、双开12.6㎡、子母2.52㎡）；门洞口四周墙面乳胶漆刷新修复1134.43㎡；木门窗拆除703.76㎡、金属门窗拆除28.35㎡'],
        ['2', '地下室车库入口挡土墙漏水整改翻新', '地面花岗石拆除及恢复200.76㎡；土方开挖、回填、余方弃置各452.24m³；室外结构墙面防水93.34㎡；室内墙面修复61.78㎡；混凝土垫层及直形墙（C30细石混凝土）'],
        ['3', '议政厅雨棚', '排水沟防水层拆除及修复182.7㎡；原张拉膜雨棚局部漏水点修复1项；钢结构基层除锈1项'],
    ],
    widths=[1.2, 3.6, 10.6],
)
para('', indent=False)
h2('1.4  建设条件')
para('项目位于单位-ZX机关院内，水、电、通信等市政配套设施齐全，交通运输条件良好，具备实施条件。施工期间可合理安排作业时段，尽量减少对机关正常办公秩序的影响。')

# 第二章
p = h1('第二章  项目现状情况及维修改造必要性')
add_bookmark(p, '_TOC_CS_2', 102)
h2('2.1  项目现状情况')
h3('2.1.1  办公用房及会议室、公共区域门现状')
para('经专业人员现场勘察，办公用房及会议室、公共区域门整体老化情况较为突出，木门结构出现老化现象，门扇框架及芯板因长期温湿度变化导致木材干缩变形、榫卯松动。部分木门插销出现松动，金属件锈蚀、固定孔扩大，无法有效锁闭，门扇内部填充材料内部空洞腐朽，敲击时声音空响、手感松软，局部门边已出现粉屑脱落。墙面存在部分破败开裂现象。')
h3('2.1.2  议政厅屋面现状')
para('经专业人员现场勘察，议政厅顶层防水层老化现象十分明显，防水层表面已呈现出不同程度、不同形态的裂缝，部分裂缝宽度较大，且裂缝深度已穿透防水层表层，失去了原有的防水阻隔作用。屋面防水的卷材防水材料出现明显脆性，质地变硬、韧性大幅降低，用手触碰易出现断裂、脱落现象，无法贴合屋面基层，防水卷材与屋面之间出现空隙、起鼓等问题。议政厅室内墙面出现大面积脱落、起鼓、返潮现象，墙面涂层斑驳脱落。')
h3('2.1.3  地下室车库入口挡土墙现状')
para('经专业人员现场勘察，办公厅负二楼出入口道路两侧墙面存在明显漏水现象。道路两侧墙面防水涂层及基层已出现老化破损，墙面裂缝、基层缝隙及墙体接缝处有水渗入。日常多为墙面返潮、局部渗水，墙面表层出现湿润斑块，用手触碰有明显水渍。降雨后，漏水现象会进一步加剧，部分区域出现水滴渗漏、墙面流水的情况，道路两侧墙面表层涂层起鼓、脱落、斑驳发黑，墙面基层出现软化、粉化现象，甚至出现细小裂缝不断扩大的情况。')
h2('2.2  维修改造必要性')
para('1.消除安全隐患、保障正常使用的需要。主楼办公用房及公共区域木门、防火门已使用23年，超出常规门窗设施10—15年的正常更换周期，门扇变形、五金件锈蚀、无法有效锁闭等问题直接影响办公安全与保密要求；议政厅屋面防水层老化失效、室内返潮脱落，负二楼出入口挡土墙漏水、基层软化粉化，长期渗漏将影响结构耐久性和正常使用。通过维修改造，可及时消除上述隐患，保障办公楼安全正常使用。')
para('2.满足正常办公功能的需要。单位-ZX主楼是机关日常办公和政务活动的重要场所，门老化、屋面渗漏、墙面漏水等问题已直接影响正常办公秩序和机关形象。实施维修改造是恢复和保障建筑功能正常发挥的必要措施。')
para('3.符合政策规定的要求。本项目属于《四川省省级机关办公用房维修管理实施办法（2023年修订版）》（川机管发〔2023〕163号）及《党政机关办公用房建设标准》（发改投资〔2014〕2674号）规定的维修项目，符合国家和四川省有关规定。')
para('4.节约投资、延长建筑寿命的需要。本项目维修内容仅限于门更换、屋面防水、挡土墙漏水整改等必要部位，不涉及主体结构改造，投资规模小、针对性强，通过适时维修可有效延长建筑整体使用寿命，避免老化问题进一步扩大造成更大的维修投入。')

# 第三章
p = h1('第三章  项目实施可行性论证')
add_bookmark(p, '_TOC_CS_3', 103)
h2('3.1  实施必要性')
para('本项目为既有办公用房维修项目，维修内容为主楼房门/防火门更换、议政厅雨棚维修、地下室车库入口挡土墙漏水整改翻新，均属《四川省省级机关办公用房维修管理实施办法（2023年修订版）》规定的维修范围，不改变原有建筑结构和使用功能，符合《党政机关办公用房建设标准》（发改投资〔2014〕2674号）关于办公用房维修的有关规定。项目实施有利于消除安全隐患、改善办公环境、保障机关正常运转，实施十分必要。')
h2('3.2  技术可行性')
para('1.维修内容均为常规维修工程，主要包括门拆除更换、屋面及墙面防水层拆除重做、地面花岗石拆除恢复、土方开挖回填等，施工工艺成熟，技术标准完善，现行国家规范和地方定额均有明确要求。')
para('2.项目不涉及主体结构改造，无特殊技术难题；门安装、防水工程、装饰装修等施工技术均已十分成熟，专业施工队伍资源充足。')
para('3.施工所需主要材料（木质门、防火门、防盗门、电子密码锁、防水卷材、防水涂料、花岗石等）市场供应充足，可就近采购，质量有保障。')
para('4.涉及既有建筑防水节点及挡土墙漏水整改的维修，严格按照《建筑与市政工程防水通用规范》GB 55030-2022等现行规范要求组织实施，确保在房屋安全的前提下进行设计施工。')
para('综上，本项目在技术上完全可行。')
h2('3.3  经济可行性')
para('本项目概算总投资154.04万元，投资规模较小。项目为维修项目，投资省、见效快，通过及时维修避免门损坏、屋面渗漏、墙面漏水等问题进一步扩大，可减少长期维修成本，具有良好的经济可行性。')
h2('3.4  社会效益及社会稳定风险')
para('项目通过维修改造改善机关办公条件、消除安全隐患、提升机关整体形象，社会效益明显。项目维修内容明确、投资规模较小，不涉及征地拆迁和移民安置，无直接利益冲突群体，施工期间可能产生的噪声、粉尘等影响将通过合理安排施工时间、采取降噪防尘措施予以控制。经评估，本项目社会稳定风险等级为低风险。')
h2('3.5  主要建设条件落实情况')
para('项目用地为机关自有用地，无需新增建设用地；水、电、通信等市政条件齐备；资金由省级财政安排解决。项目立项、规划等前置条件落实，具备实施条件。')

# 第四章
p = h1('第四章  维修设计说明')
add_bookmark(p, '_TOC_CS_4', 104)
h2('4.1  设计原则')
para('1.安全可靠原则：在确保房屋结构安全的前提下进行设计，涉及既有建筑防水节点及挡土墙漏水整改的，严格按照现行强制性规范要求执行。')
para('2.经济适用原则：严格按照批准的维修规划内容和《党政机关办公用房建设标准》（发改投资〔2014〕2674号）要求设计，不超标准、不扩大维修范围，合理控制投资。')
para('3.节能环保原则：选用节能环保材料，施工过程采取降噪防尘措施，建筑垃圾及时清运处置。')
para('4.功能恢复原则：以恢复和改善办公用房使用功能为目标，维修内容与既有建筑现状相匹配，注重耐久性和实用性。')
h2('4.2  设计依据')
norms = [
    '《四川省政府投资管理办法》（川府规〔2022〕5号）',
    '《政府投资项目可行性研究报告编写通用大纲》（2023年版）',
    '《党政机关办公用房建设标准》（发改投资〔2014〕2674号）',
    '《四川省省级机关办公用房维修管理实施办法（2023年修订版）》（川机管发〔2023〕163号）',
    '《建筑工程设计文件编制深度规定》（2016年版）',
    '《建筑与市政工程防水通用规范》GB 55030-2022',
    '《屋面工程技术规范》GB 50345-2012',
    '《屋面工程质量验收规范》GB 50207-2012',
    '《建筑装饰装修工程质量验收标准》GB 50210-2018',
    '《建筑内部装修设计防火规范》GB 50222-2017',
    '《公共建筑节能设计标准》GB 50189-2015',
    '《建设工程工程量清单计价规范》GB 50500-2013',
    '《房屋建筑与装饰工程工程量计算规范》GB 50854-2013',
    '《四川省建设工程工程量清单计价定额》（2020）',
    '国家和四川省其他现行有关工程建设标准、规范及规定',
]
for i, nm in enumerate(norms, 1):
    para(f'（{i}）{nm}。', size=12)
h2('4.3  主要维修改造内容')
h3('4.3.1  主楼房门/防火门更换')
para('对办公区域所有原有木门（含会议室老化木门）及防火门进行拆除，选用耐用性强的木门及防火门。拆除木门窗703.76㎡、金属门窗28.35㎡，新装门锁（电子密码锁）130个、木质门带套（单开466.83㎡、双开81.32㎡）、防火门（单开乙级65.52㎡、双开乙级88.2㎡、单开甲级1.89㎡）、防盗门（单开13.23㎡、双开12.6㎡、子母2.52㎡）；对门洞口四周墙面进行乳胶漆刷新修复1134.43㎡，对开裂墙面开裂处周边松动的墙面涂层、浮灰、杂物清理干净，去除内部腐朽、松动的基层，确保修复基层牢固，选用环保、无异味、耐擦洗、不易褪色的优质乳胶漆，与原有墙面乳胶漆颜色保持一致。')
h3('4.3.2  地下室车库入口挡土墙漏水整改翻新')
para('对办公厅负二楼出入口道路两侧墙面进行全面排查，明确漏水点位置、漏水程度及漏水来源。拆除地面花岗石200.76㎡，土方开挖、回填、余方弃置各452.24m³；对漏水区域墙面进行清理，铲除起鼓、脱落的涂层及软化、粉化的基层，直至露出坚实的墙体基层。选用适配地下潮湿环境、防水防潮性能优良、耐霉变、附着力强的专用防水涂层材料，对室外结构墙面做防水93.34㎡（20mm厚抗裂水泥砂浆抹灰层内设耐碱玻纤网格布、SBS改性沥青卷材4mm+3mm双层、20mm厚抗裂砂浆保护层）；室内墙面修复61.78㎡（界面剂1道、涂刷丙烯酸弹性防水涂料2道1.2mm厚、5mm厚抗裂砂浆保护层、外墙耐水腻子2遍、外墙抗碱封闭底漆1遍、外墙白色面漆2遍）；地面花岗石恢复200.76㎡（50mm厚花岗石饰面、10mm素水泥浆粘接层、50mm干硬性水泥砂浆结合层、200mm厚5%水泥稳定碎石基层、300mm厚级配碎石垫层），并施工C30细石混凝土垫层及直形墙。')
h3('4.3.3  议政厅雨棚维修')
para('对议政厅老化破损的原有防水层及脆性卷材进行拆除，拆除排水沟防水层182.7㎡，拆除过程中采用专业工具，避免损坏屋面基层结构，拆除的废旧卷材、防水材料等建筑垃圾及时清理清运。新做排水沟防水修复182.7㎡（1.5mm厚高强水性橡胶沥青防水涂料、1.5mm厚纤维增强型高分子膜基自粘沥青防水卷材、30mm厚1:6水泥砂浆2%找坡层、干铺无纺聚酯纤维布隔离层、20mm厚C20细石混凝土）；针对原张拉膜雨棚局部漏水点进行修复1项，钢结构基层除锈1项。')

h2('4.4  主要材料或设备表')
add_table(
    ['序号', '材料（设备）名称', '规格型号/技术要求', '单位', '数量', '备注'],
    [
        ['1', '电子密码锁', '详设计', '把', '130', '门更换'],
        ['2', '木质门带套（单开）', '灰色木纹，含五金锁具', '㎡', '466.83', '门更换'],
        ['3', '木质门带套（双开）', '灰色木纹，含五金锁具', '㎡', '81.32', '门更换'],
        ['4', '单开乙级防火门', '含闭门器', '㎡', '65.52', '门更换'],
        ['5', '双开乙级防火门', '含闭门器', '㎡', '88.2', '门更换'],
        ['6', '单开甲级防火门', '含闭门器', '㎡', '1.89', '门更换'],
        ['7', '单开/双开/子母防盗门', 'FDM0921/FDM1521/FDM1221', '㎡', '28.35', '门更换'],
        ['8', 'SBS改性沥青防水卷材', '4mm+3mm双层', '㎡', '93.34', '挡土墙防水'],
        ['9', '丙烯酸弹性防水涂料', '涂刷2道，厚1.2mm', '㎡', '61.78', '室内墙面修复'],
        ['10', '纤维增强型高分子膜基自粘沥青防水卷材', '1.5mm厚', '㎡', '182.7', '议政厅排水沟'],
        ['11', '高强水性橡胶沥青防水涂料', '1.5mm厚', '㎡', '182.7', '议政厅排水沟'],
        ['12', '花岗石板', '厚50mm', '㎡', '200.76', '地面恢复'],
        ['13', '商品细石混凝土C30', '详设计', 'm³', '约37', '垫层及直形墙'],
        ['14', '白色乳胶漆（底漆/面漆）', '环保型，一底两面', '㎡', '1134.43', '门洞口四周墙面'],
    ],
    widths=[1.2, 4.2, 4.4, 1.5, 1.6, 2.5],
)
para('', indent=False)
h2('4.5  设计内容及做法')
h3('4.5.1  门更换做法')
para('1.拆除原有老化木门、防火门及金属门窗，拆除时做好成品保护及安全防护，建渣及时清运；')
para('2.门安装：新装电子密码锁130个、木质门带套、防火门、防盗门等，门框与墙体间缝隙填塞密实，五金配件齐全，启闭灵活；')
para('3.门洞口四周墙面乳胶漆刷新修复：清理开裂墙面松动的涂层、浮灰、杂物，去除内部腐朽、松动的基层，涂刷环保优质乳胶漆一底两面，与原有墙面颜色保持一致。')
h3('4.5.2  挡土墙漏水整改翻新做法')
para('1.拆除地面花岗石200.76㎡，土方开挖、回填、余方弃置各452.24m³；')
para('2.室外结构墙面防水：20mm厚抗裂水泥砂浆抹灰层内设耐碱玻纤网格布，铺贴SBS改性沥青卷材（4mm+3mm双层），外做20mm厚抗裂砂浆保护层；')
para('3.室内墙面修复：界面剂1道，涂刷丙烯酸弹性防水涂料2道（1.2mm厚），5mm厚抗裂砂浆保护层，外墙耐水腻子2遍，外墙抗碱封闭底漆1遍，外墙白色面漆2遍；')
para('4.地面花岗石恢复：50mm厚花岗石饰面（5mm留缝）、10mm素水泥浆粘接层、50mm干硬性水泥砂浆结合层、200mm厚5%水泥稳定碎石基层、300mm厚级配碎石垫层；')
para('5.施工C30细石混凝土垫层及直形墙，满足设计及规范要求。')
h3('4.5.3  议政厅雨棚维修做法')
para('1.拆除排水沟原有防水层182.7㎡，清理基层；')
para('2.新做排水沟防水修复：1.5mm厚高强水性橡胶沥青防水涂料，1.5mm厚纤维增强型高分子膜基自粘沥青防水卷材，30mm厚1:6水泥砂浆2%找坡层，干铺无纺聚酯纤维布隔离层，20mm厚C20细石混凝土；')
para('3.针对原张拉膜雨棚局部漏水点进行修复，钢结构基层除锈，确保雨棚防水及结构完好。')

# 第五章
p = h1('第五章  安全设施设计')
add_bookmark(p, '_TOC_CS_5', 105)
h2('5.1  结构安全')
para('本项目为维修项目，不改变原建筑主体结构体系。地下室车库入口挡土墙漏水整改涉及土方开挖（暂定深度3800mm，实际深度依据现场结构确定）及地面结构恢复，严格按照《建筑与市政工程防水通用规范》GB 55030-2022等现行强制性规范要求实施，施工前对既有结构进行现场查勘，明确结构受力状态，确保房屋整体安全。')
h2('5.2  施工安全')
para('1.施工前编制专项施工方案，落实安全生产责任制，对作业人员进行安全技术交底；')
para('2.拆除及土方开挖作业设置警戒区域，做好临边防护；屋面及高空作业搭设符合要求的脚手架并系挂安全带，设置防坠落措施；')
para('3.施工现场用电由专业电工负责，严格执行临时用电安全规范，做到"一机一闸一漏一箱"；')
para('4.施工机械进场前检验合格，操作人员持证上岗；建筑垃圾吊运采用专用容器，严禁高空抛掷。')
h2('5.3  消防安全')
para('1.防水卷材热熔铺贴等动火作业严格执行动火审批制度，配备灭火器材，作业后落实看火措施；')
para('2.施工期间保持楼梯间、疏散通道畅通，不占用、不堵塞消防设施，不改变原建筑消防分区和疏散体系，恢复施工破坏的消防设施；')
para('3.更换的防火门为甲级、乙级防火门，满足相应耐火极限要求。')
h2('5.4  电气及给排水安全')
para('1.施工前办理停电手续，施工过程中不得损坏既有电气线路及设备，如发现损坏及时修复；')
para('2.临时用电线路敷设规范，配电箱设置漏电保护装置；')
para('3.涉及防水及土方作业时，做好现场排水和临边防护，确保施工安全。')

# 第六章
p = h1('第六章  节能设计')
add_bookmark(p, '_TOC_CS_6', 106)
h2('6.1  节能原则')
para('本项目为维修项目，在满足使用功能和现行标准的前提下，通过选用节能环保材料、恢复屋面防水保温性能、更换节能门等措施，实现节能降耗目标。')
h2('6.2  节能措施')
para('1.更换老化、变形、密封失效的木门及防火门，新门密封性能良好，减少冷热风渗透，降低建筑能耗；')
para('2.恢复议政厅屋面及挡土墙防水性能，减少因渗漏返潮导致的室内能耗损失；')
para('3.选用符合国家节能环保标准的建材，室内装修材料甲醛、挥发性有机物等有害物质含量满足国家限量标准；')
para('4.施工期间合理安排作业时段，减少能源消耗。')

# 第七章
p = h1('第七章  环境保护及影响评价')
add_bookmark(p, '_TOC_CS_7', 107)
h2('7.1  环境影响分析')
para('本项目施工期主要环境影响包括：拆除及土方作业产生的噪声、粉尘，建筑垃圾及废弃材料（废旧门窗、防水材料、花岗石、土方等）的运输处置，少量施工废水等。项目位于单位-ZX机关院内，周边为办公区域，施工期环境影响主要集中在施工阶段，采取有效措施后可控制在可接受范围内。')
h2('7.2  污染防治措施')
para('1.噪声控制：合理安排施工作业时间，避开机关办公时段；选用低噪声设备，高噪声作业采取隔声降噪措施；')
para('2.粉尘控制：施工区域设置围挡，拆除作业采用湿法作业，材料堆放覆盖，运输车辆密闭，出入口设置冲洗设施；')
para('3.废水控制：施工废水经沉淀处理后排放，不得随意排入市政管网；')
para('4.建筑垃圾处置：拆除的废旧门窗、防水卷材、花岗石、土方等建筑垃圾分类收集、及时清运，按照《四川省住房和城乡建设厅关于规范建设工程建筑垃圾和拆除垃圾外运及处置费计价的通知》（川建标函〔2024〕3160号）等规定进行外运及处置，保持施工区域整洁，减少对周边环境及机关正常办公的影响；')
para('5.材料环保：选用符合国家环保标准的绿色建材，施工完成后对室内环境进行检测，确保空气质量达标。')

# 第八章
p = h1('第八章  项目实施组织管理')
add_bookmark(p, '_TOC_CS_8', 108)
h2('8.1  组织机构')
para('项目由单位-ZX组织实施，成立项目建设管理机构，明确专人负责项目建设管理，落实项目法人责任制，对项目建设的进度、质量、投资、安全实行全过程管理。')
h2('8.2  进度安排')
para('项目计划工期180日历天。拟按以下阶段组织实施：施工准备（含招标、合同签订、图纸会审）约30日历天；主体维修施工（门更换、挡土墙漏水整改翻新、议政厅雨棚维修）约120日历天；竣工验收及结算约30日历天。施工期间结合机关办公实际，合理安排作业时段，确保按期完工。')
h2('8.3  招标方案')
para('按照《中华人民共和国招标投标法》及四川省有关规定，本项目施工拟采用公开招标方式，委托具有相应资质的招标代理机构组织招标，招标范围为全部工程施工内容，招标组织形式为委托招标。设计、监理等服务按国家有关规定确定。招标方案如下表：')
add_table(
    ['招标内容', '招标范围', '招标组织形式', '招标方式', '估算金额（万元）'],
    [
        ['建筑工程施工', '全部招标', '委托招标', '公开招标', '132.59'],
        ['工程监理', '全部招标', '委托招标', '邀请招标或竞争性方式', '3.72'],
        ['设计', '全部招标', '委托招标', '竞争性方式', '5.07'],
    ],
    widths=[3.8, 2.4, 2.4, 3.0, 3.8],
)
para('', indent=False)
h2('8.4  合同管理')
para('严格执行合同管理制度，依法签订设计、施工、监理等合同，明确各方权利义务、质量标准、工期要求和违约责任；加强合同履约管理，规范工程变更签证程序。')
h2('8.5  质量与投资控制')
para('1.建立工程质量保证体系，实行工程监理制，关键工序、隐蔽工程（防水层、土方回填等）实行旁站和验收制度，确保工程质量符合设计和规范要求；')
para('2.严格按照批准的投资概算控制工程投资，加强工程计量与支付管理，防止超概算；')
para('3.工程竣工后及时组织竣工验收，办理竣工结算和财务决算，接受审计监督。')

# 第九章
p = h1('第九章  投资概算与资金筹措')
add_bookmark(p, '_TOC_CS_9', 109)
h2('9.1  概算编制依据')
para('1.单位-ZX主楼综合维修改造项目设计图纸；')
para('2.《建设工程工程量清单计价规范》GB 50500-2013及《房屋建筑与装饰工程工程量计算规范》GB 50854-2013等现行工程量计算规范；')
para('3.《四川省建设工程工程量清单计价定额》（2020）及相关配套和解释文件；')
para('4.《四川省住房和城乡建设厅关于重新调整<建筑业营业税改征增值税四川省建设工程计价依据调整办法>的通知》（川建造价发〔2019〕181号）；')
para('5.人工费调整执行川建价发〔2025〕41号《关于批准2026年上半年各市（州）2020年〈四川省建设工程工程量清单计价定额〉人工费调整幅度及计日工人工单价的批复》；')
para('6.安全文明施工费执行《四川省建设工程安全文明施工费计价管理办法》（川建行规〔2024〕15号），其他总价措施取费标准执行川建标函〔2024〕3159号；')
para('7.材料价格采用2026年第6期《建设工程价格信息》成都市市区价格，无对应价格的采用市场询价；')
para('8.规费费率按照I档计取，销项增值税税率按9%计取，附加税税率按0.313%计取；')
para('9.《基本建设项目建设成本管理规定》（财建〔2016〕504号）、《建设工程监理与相关服务收费管理规定》（发改价格〔2007〕670号）、《工程勘察设计收费管理规定》（计价格〔2002〕10号）、《国家发展改革委关于降低部分建设项目收费标准规范收费行为等有关问题的通知》（川发改价格〔2011〕323号）、《招标代理服务收费管理暂行办法》（计价格〔2002〕1980号）、《工程造价咨询服务收费标准》（川价发〔2008〕141号）等有关取费规定。')
h2('9.2  概算范围')
para('本项目投资概算范围为经批准维修内容所包含的工程费用、工程建设其他费用和预备费，即：主楼房门/防火门更换、地下室车库入口挡土墙漏水整改翻新、议政厅雨棚维修等全部维修内容。')
h2('9.3  项目投资概算')
para('项目概算总投资154.04万元，其中：第一部分建筑安装工程费用132.59万元，第二部分工程建设其他费用14.12万元，第三部分预备费7.34万元（基本预备费，按第一、二部分费用之和的5%计取）。投资概算汇总如下表：')
add_table(
    ['序号', '费用名称', '概算金额（万元）', '计算说明'],
    [
        ['一', '第一部分  建筑安装工程费用', '132.59', '其中主楼房门/防火门更换67.59万元、挡土墙漏水整改翻新22.02万元、议政厅雨棚10.54万元等'],
        ['二', '第二部分  工程建设其他费用', '14.12', ''],
        ['', '其中：建设单位管理费', '2.65', '参照财建〔2016〕504号'],
        ['', '工程监理费', '3.72', '参照发改价格〔2007〕670号'],
        ['', '设计费', '5.07', '参照计价格〔2002〕10号'],
        ['', '设计文件审查费', '0.27', '参照川发改价格〔2011〕323号'],
        ['', '工程招标代理费', '1.23', '参照计价格〔2002〕1980号'],
        ['', '招标工程量清单及控制价编制费', '0.52', '参照川价发〔2008〕141号'],
        ['', '竣工结算审核费', '0.66', '参照川价发〔2008〕141号'],
        ['三', '第三部分  预备费', '7.34', '按第一、二部分费用之和的5%'],
        ['四', '概算总投资', '154.04', ''],
    ],
    widths=[1.2, 6.2, 3.0, 5.0],
)
para('', indent=False)
para('建筑安装工程费用中，各分部分项工程费用构成如下表：')
add_table(
    ['序号', '分部工程名称', '金额（万元）', '主要内容'],
    [
        ['1', '主楼房门/防火门更换', '67.59', '门锁130个、木质门、防火门、防盗门更换及门洞口四周墙面乳胶漆修复'],
        ['2', '地下室车库入口挡土墙漏水整改翻新', '22.02', '地面花岗石拆除恢复、土方开挖回填、室外结构墙面防水、室内墙面修复'],
        ['3', '议政厅雨棚', '10.54', '排水沟防水层拆除修复、张拉膜雨棚漏水点修复、钢结构基层除锈'],
        ['4', '单价措施项目', '4.55', '脚手架、混凝土模板及支架、垂直运输、大型机械设备进出场及安拆'],
        ['', '分部分项及单价措施项目小计', '104.70', ''],
        ['5', '总价措施项目', '7.27', '含安全文明施工费7.02万元'],
        ['6', '其他项目（暂列金额）', '7.84', ''],
        ['7', '规费', '1.48', '按I档计取'],
        ['8', '税金', '11.30', '销项增值税及附加税'],
        ['', '建筑安装工程费用合计', '132.59', ''],
    ],
    widths=[1.2, 5.0, 3.2, 6.0],
)
para('', indent=False)
para('注：上表尾数按四舍五入保留两位小数；工程垃圾、拆除垃圾外运及处置费按川建标函〔2024〕3160号规定由投标人自主报价、包干使用，结算时不作调整。')
h2('9.4  投资控制')
para('本项目概算总投资154.04万元，项目在实施过程中将严格执行批准的投资概算，加强投资控制，确保不突破概算。')
h2('9.5  资金筹措')
para('项目资金来源为省级财政资金，由项目单位按程序申请拨付。项目概算总投资154.04万元，资金需求纳入年度预算安排，确保项目顺利实施。')

# 第十章
p = h1('第十章  结论与建议')
add_bookmark(p, '_TOC_CS_10', 110)
h2('10.1  结论')
para('1.单位-ZX主楼综合维修改造项目（办公用房维修项目）维修内容符合《党政机关办公用房建设标准》（发改投资〔2014〕2674号）及《四川省省级机关办公用房维修管理实施办法（2023年修订版）》有关规定。')
para('2.项目针对主楼门老化、议政厅屋面渗漏、地下室车库入口挡土墙漏水等突出问题实施维修，技术可行、经济合理，实施必要。')
para('3.项目概算总投资154.04万元，资金来源为省级财政资金，资金有保障。')
para('4.项目实施不影响主体结构安全，施工组织条件具备，社会稳定风险低。')
para('综上所述，项目具备实施条件，建议批准实施。')
h2('10.2  建议')
para('1.建议尽快按程序审批项目初步设计和投资概算，批复后及时开展施工招标，落实建设资金，确保项目按期实施；')
para('2.施工期间合理安排作业时段，尽量减少对机关正常办公秩序的影响；')
para('3.严格执行批准的投资概算，加强全过程造价控制，确保项目投资控制在概算以内；')
para('4.严格落实安全生产和环境保护各项措施，确保施工安全和文明施工。')

# ===================== TOC1 样式 =====================
try:
    toc1 = doc.styles['TOC 1']
except KeyError:
    toc1 = doc.styles.add_style('TOC 1', WD_STYLE_TYPE.PARAGRAPH)
toc1.font.name = FONT
toc1.font.size = Pt(12)
if toc1.element.rPr is None:
    rPr1 = OxmlElement('w:rPr')
    toc1.element.append(rPr1)
toc1.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
toc1.paragraph_format.line_spacing = Pt(24)
toc1.paragraph_format.tab_stops.add_tab_stop(Cm(15.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

# ===================== 页脚页码（节3） =====================
sec3 = doc.sections[2]
pgNumType = OxmlElement('w:pgNumType')
pgNumType.set(qn('w:fmt'), 'decimal')
pgNumType.set(qn('w:start'), '1')
sec3._sectPr.append(pgNumType)
sec3.footer.is_linked_to_previous = False
footer = sec3.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
fp.paragraph_format.line_spacing = Pt(16)
r = fp.add_run('第 '); set_run_font(r, size=10.5)
add_field(fp, ' PAGE ', '1')
r = fp.add_run(' 页/共 '); set_run_font(r, size=10.5)
add_field(fp, ' SECTIONPAGES ', '1')
r = fp.add_run(' 页'); set_run_font(r, size=10.5)


def remove_front_footer_refs():
    for s in doc.sections[:2]:
        sectPr = s._sectPr
        for fr in sectPr.findall(qn('w:footerReference')):
            sectPr.remove(fr)


settings = doc.settings.element
uf = OxmlElement('w:updateFields')
uf.set(qn('w:val'), 'true')
settings.append(uf)


def fix_all_fonts():
    from docx.oxml.ns import qn as _qn
    def fix_run(run):
        rPr = run._element.get_or_add_rPr()
        rf = rPr.find(_qn('w:rFonts'))
        if rf is None:
            rf = OxmlElement('w:rFonts')
            rPr.insert(0, rf)
        for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
            rf.set(_qn(a), FONT)
    for para_ in doc.paragraphs:
        for run in para_.runs:
            fix_run(run)
        for hl in para_._p.findall(qn('w:hyperlink')):
            for r_ in hl.findall(qn('w:r')):
                rPr = r_.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr')
                    r_.insert(0, rPr)
                rf = rPr.find(qn('w:rFonts'))
                if rf is None:
                    rf = OxmlElement('w:rFonts')
                    rPr.insert(0, rf)
                for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
                    rf.set(qn(a), FONT)
    for t_ in doc.tables:
        for row in t_.rows:
            for cell in row.cells:
                for para_ in cell.paragraphs:
                    for run in para_.runs:
                        fix_run(run)
    for footer_ in doc.sections[2].footer.paragraphs:
        for run in footer_.runs:
            fix_run(run)


fix_all_fonts()
remove_front_footer_refs()

doc.save(PATH)
print('SAVED:', PATH)
print('段落数:', len(doc.paragraphs), '表格数:', len(doc.tables), '节数:', len(doc.sections))
