# -*- coding: utf-8 -*-
"""项目-LZ办公区综合维修改造项目 办公用房维修项目初步设计（代项目建议书、可行性研究报告）生成脚本
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

PATH = r'<桌面>\改造项目PDF图纸\改造项目PDF图纸\项目-LZ办公区综合维修改造项目\项目-LZ办公区综合维修改造项目初步设计（代项目建议书、可行性研究报告）.docx'
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
para('项目-LZ办公区综合维修改造项目', size=26, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=40)
para('', indent=False)
para('初步设计', size=26, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=40)
para('（代项目建议书、可行性研究报告）', size=22, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=36)
for _ in range(11):
    para('', indent=False)
para('项目单位：单位-LZ', size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=34)
para('建设地点：地址-LZ', size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=34)
para_design_unit('设计单位：' + '\u3000' * 16)
para('编制时间：二〇二六年八月', size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, line=34)

sec1 = doc.add_section(WD_SECTION.NEW_PAGE)
sec1.page_width = Cm(21.0); sec1.page_height = Cm(29.7)
sec1.top_margin = Cm(3.0); sec1.bottom_margin = Cm(2.5)
sec1.left_margin = Cm(2.8); sec1.right_margin = Cm(2.6)

# ===================== 编制说明（节2） =====================
heading('编 制 说 明', size=22, before=6, after=12)
bs_items = [
    '一、本初步设计（代项目建议书、可行性研究报告）依据《四川省政府投资管理办法》（川府规〔2022〕5号）规定，按照"三合一"评审（项目建议书、可行性研究报告、初步设计）要求编制，用于项目-LZ办公区综合维修改造项目办公用房维修项目的审批。',
    '二、项目单位：单位-LZ。项目建设地点位于地址-LZ。',
    '三、本文件主要内容包括：项目基本概况，项目现状情况及维修改造必要性，项目实施可行性论证，维修设计说明，安全设施设计、节能设计，环境保护及影响评价，项目实施组织管理，投资概算与资金筹措等。',
    '四、编制依据：《四川省政府投资管理办法》（川府规〔2022〕5号）、《政府投资项目可行性研究报告编写通用大纲》（2023年版）、《建筑工程设计文件编制深度规定》（2016年版）、《四川省省级机关办公用房维修管理实施办法（2023年修订版）》（川机管发〔2023〕163号）、《党政机关办公用房建设标准》（发改投资〔2014〕2674号）以及国家、四川省现行有关工程建设标准、规范及规定。',
    '五、项目投资概算依据设计图纸、现行工程量清单计价规范和四川省建设工程工程量清单计价定额等编制，项目概算总投资318.12万元。',
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
        ['项目名称', '项目-LZ办公区综合维修改造项目（办公用房维修项目）'],
        ['项目单位', '单位-LZ'],
        ['建设地点', '地址-LZ'],
        ['项目性质', '党政机关办公用房维修项目'],
        ['维修内容', '2号楼多联机空调系统更新、室外训练场改造、3号楼车位及相邻草坪硬化'],
        ['建设工期', '180日历天'],
        ['投资概算', '项目概算总投资318.12万元'],
        ['资金来源', '申请省级财政资金'],
    ],
    widths=[3.0, 12.4],
)
para('', indent=False)
h2('1.2  建筑基本情况')
add_table(
    ['项目', '内容'],
    [
        ['建筑名称', '单位-LZ1号楼、2号楼、3号楼'],
        ['结构形式', '框架结构'],
        ['建成年代', '2号楼空调系统于2013年安装'],
        ['已使用年限', '——'],
        ['使用性质', '纪检监察业务用房'],
        ['坐落位置', '地址-LZ'],
        ['维修范围', '2号楼空调系统、室外训练场、3号楼车位及相邻草坪等部位'],
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
        ['1', '2号楼多联机空调系统更新', '采购并安装全新多联机空调设备，直流变频室外机5种型号（GMV-450WM/S、GMV-400WM/S、GMV-615WM/S、GMV-785WM/S、GMV-900WM/S）共12台；原主机拆除吊装、原室内机清洗检测；配套线控器、连接铜管（φ6.4—28.6mm多规格）及单/双层百叶风口等'],
        ['2', '室外训练场改造', '室外训练场塑化面层全面重新铺装约3000㎡，含足球场拆除（面积大于5000㎡）、人工草皮、塑胶跑道、拼装弹性运动地板；对原破损塑化面层拆除、基层平整压实'],
        ['3', '3号楼车位及相邻草坪硬化', '3号楼现有地面机动车位26个及相邻草坪区域统一硬化，拆除草坪砖及绿植隔离带，级配碎石垫层、沥青混凝土面层；移建充电桩6个、路灯5个、监控2个'],
    ],
    widths=[1.2, 3.6, 10.6],
)
para('', indent=False)
h2('1.4  建设条件')
para('项目位于单位-LZ院内，水、电、通信等市政配套设施齐全，交通运输条件良好，具备实施条件。施工期间可合理安排作业时段，尽量减少对机关正常办公秩序的影响。')

# 第二章
p = h1('第二章  项目现状情况及维修改造必要性')
add_bookmark(p, '_TOC_CS_2', 102)
h2('2.1  项目现状情况')
h3('2.1.1  2号楼空调系统现状')
para('经专业人员现场勘察，2号楼多联机空调主机安装位置通风散热条件差、设备老化，现频繁出现内外机故障、房间制冷效果差、冷凝水漏水等问题。近两年来已多次出现故障后无法及时修复的情况，较大程度影响了暖气、冷气的正常供应，已难以满足纪检监察业务用房正常办公使用需求。')
h3('2.1.2  室外训练场现状')
para('经专业人员现场勘察，室外训练场面积约4100㎡，包含沙坑（约300㎡）、无障碍训练设施、跑道、篮球场、羽毛球场及草坪。因看护人数多、训练强度大、使用频率高，现室外训练场塑胶铺装区域已有一定程度损坏，部分区域出现开裂、起皮、空鼓等现象，影响正常训练使用。')
h3('2.1.3  3号楼车位及相邻草坪现状')
para('经专业人员现场勘察，3号楼区域现有地面机动车位26个。车位相邻区域存在部分草坪，但草坪与车位之间以及车位与园区道路中间均设置了绿植隔离带。由于隔离带的物理阻隔，车辆进出不便、视线受阻，部分车位实际使用率较低，造成车位闲置，园区通行与停车组织不便。')
h2('2.2  维修改造必要性')
para('1.消除安全隐患、保障正常使用的需要。2号楼多联机空调主机通风散热差、设备老化，频繁出现内外机故障、制冷效果差、冷凝水漏水等问题，已较大程度影响暖冷气正常供应；室外训练场塑胶铺装损坏影响正常训练使用；3号楼车位及草坪区域绿植隔离带阻碍车辆进出、视线受阻。通过维修改造，可及时消除上述隐患，保障办案服务中心安全正常使用。')
para('2.满足正常办公功能的需要。单位-LZ是纪检监察业务用房，空调系统失效、训练场损坏、车位闲置等问题已直接影响正常办公、训练和停车组织。实施维修改造是恢复和保障建筑功能正常发挥的必要措施。')
para('3.符合政策规定的要求。本项目属于《四川省省级机关办公用房维修管理实施办法（2023年修订版）》（川机管发〔2023〕163号）及《党政机关办公用房建设标准》（发改投资〔2014〕2674号）规定的维修项目，符合国家和四川省有关规定。')
para('4.节约投资、延长设施寿命的需要。本项目维修内容仅限于空调系统更新、训练场铺装恢复、车位及草坪硬化等必要部位，不涉及主体结构改造，投资规模小、针对性强，通过适时维修可有效延长既有设施使用寿命，避免老化问题进一步扩大造成更大的维修投入。')

# 第三章
p = h1('第三章  项目实施可行性论证')
add_bookmark(p, '_TOC_CS_3', 103)
h2('3.1  实施必要性')
para('本项目为既有办公用房维修项目，维修内容为2号楼多联机空调系统更新、室外训练场改造、3号楼车位及相邻草坪硬化，均属《四川省省级机关办公用房维修管理实施办法（2023年修订版）》规定的维修范围，不改变原有建筑结构和使用功能，符合《党政机关办公用房建设标准》（发改投资〔2014〕2674号）关于办公用房维修的有关规定。项目实施有利于消除安全隐患、改善办公条件、保障机关正常运转，实施十分必要。')
h2('3.2  技术可行性')
para('1.维修内容均为常规维修工程，主要包括多联机空调设备采购安装、塑胶场地铺装、场地硬化等，施工工艺成熟，技术标准完善，现行国家规范和地方定额均有明确要求。')
para('2.项目不涉及主体结构改造，无特殊技术难题；空调安装、塑胶铺装、沥青混凝土路面等施工技术均已十分成熟，专业施工队伍资源充足。')
para('3.施工所需主要材料（多联机空调设备、人工草皮、塑胶跑道材料、拼装运动地板、沥青混凝土、级配碎石等）市场供应充足，可就近采购，质量有保障。')
para('4.涉及既有建筑空调系统更换及室外场地改造的维修，严格按照现行规范要求组织实施，确保在房屋安全的前提下进行设计施工。')
para('综上，本项目在技术上完全可行。')
h2('3.3  经济可行性')
para('本项目概算总投资318.12万元，投资规模较小。项目为维修项目，投资省、见效快，通过及时维修避免空调系统老化失效、训练场损坏、车位闲置等问题进一步扩大，可减少长期维修成本，具有良好的经济可行性。')
h2('3.4  社会效益及社会稳定风险')
para('项目通过维修改造改善机关办公条件、消除安全隐患、提升机关整体形象，社会效益明显。项目维修内容明确、投资规模较小，不涉及征地拆迁和移民安置，无直接利益冲突群体，施工期间可能产生的噪声、粉尘等影响将通过合理安排施工时间、采取降噪防尘措施予以控制。经评估，本项目社会稳定风险等级为低风险。')
h2('3.5  主要建设条件落实情况')
para('项目用地为机关自有用地，无需新增建设用地；水、电、通信等市政条件齐备；资金由省级财政安排解决。项目立项、规划等前置条件落实，具备实施条件。')

# 第四章
p = h1('第四章  维修设计说明')
add_bookmark(p, '_TOC_CS_4', 104)
h2('4.1  设计原则')
para('1.安全可靠原则：在确保房屋结构安全的前提下进行设计，涉及既有建筑空调系统更换及室外场地改造的，严格按照现行强制性规范要求执行。')
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
h3('4.3.1  2号楼多联机空调系统更新')
para('对2号楼老化多联机空调系统进行更新，开展全新空调设备采购与安装工作。按照2号楼实际制冷制热需求，采购直流变频多联机空调设备，其中室外机共12台（含GMV-450WM/S、GMV-400WM/S、GMV-615WM/S、GMV-785WM/S、GMV-900WM/S共5种型号），并采购全新连接管线、控制设备及附属配件。拆除原有老化主机并吊装清运，对原有室内机进行清洗、检测维修；精准铺设连接铜管、调试控制线路，确保内外机运行协调，制冷制热效果达标，杜绝冷凝水漏水。')
h3('4.3.2  室外训练场改造')
para('对室外训练场塑化部分进行全面重新铺装，铺装范围涵盖无障碍训练设施带、跑道、篮球场、羽毛球场等，总面积约3000㎡。对原有破损塑化面层进行拆除，清理场地基层，去除表面破损材料、杂物及松动土壤，对基层进行平整、压实处理，确保基层坚实平整。重新铺设足球场人工草皮、塑胶跑道及拼装弹性运动地板，满足训练使用要求。')
h3('4.3.3  3号楼车位及相邻草坪硬化')
para('将3号楼现有的26个地面机动车位及相邻草坪区域进行统一硬化处理。拆除原有绿植隔离带，平整场地后铺设高强度混凝土或沥青面层，形成连续、平整的大面积硬化场地。移建该区域充电桩6个、路灯5个、监控2个。')

h2('4.4  主要材料或设备表')
add_table(
    ['序号', '材料（设备）名称', '规格型号/技术要求', '单位', '数量', '备注'],
    [
        ['1', '直流变频室外机', 'GMV-450WM/S', '台', '2', '空调系统'],
        ['2', '直流变频室外机', 'GMV-400WM/S', '台', '2', '空调系统'],
        ['3', '直流变频室外机', 'GMV-615WM/S', '台', '4', '空调系统'],
        ['4', '直流变频室外机', 'GMV-785WM/S', '台', '2', '空调系统'],
        ['5', '直流变频室外机', 'GMV-900WM/S', '台', '2', '空调系统'],
        ['6', '线控器', '配套多联机空调', '个', '187', '空调系统'],
        ['7', '连接铜管', 'φ6.4—28.6mm多规格', '项', '1', '空调系统'],
        ['8', '单/双层百叶风口', '配套', '项', '1', '空调系统'],
        ['9', '人工草皮', '足球场专用', '㎡', '约1200', '训练场'],
        ['10', '塑胶跑道材料', '现浇型塑胶面层', '㎡', '约2500', '训练场'],
        ['11', '拼装弹性运动地板', '户外拼装型', '㎡', '约600', '训练场'],
        ['12', '沥青混凝土', 'AC-20中粒式/AC-13细粒式', '㎡', '644', '车位硬化'],
        ['13', '级配碎石', '500mm厚分层压实', 'm³', '386.4', '车位硬化'],
        ['14', '水泥稳定碎石', '5%水泥稳定，200mm厚', 'm³', '193.2', '车位硬化'],
    ],
    widths=[1.2, 4.2, 4.4, 1.5, 1.6, 2.5],
)
para('', indent=False)
h2('4.5  设计内容及做法')
h3('4.5.1  空调系统更新做法')
para('1.拆除原有老化室外主机及连接管线，拆除设备吊运至建设单位指定位置，做好成品保护及安全防护；')
para('2.新装直流变频室外机12台，主机基础加固处理，新主机吊装就位，安装线控器187个，铺设连接铜管（φ6.4—28.6mm多规格）；')
para('3.原有室内机清洗、检测维修，原系统管道冲洗，系统检查、调试，确保内外机运行协调、制冷制热达标、无冷凝水漏水。')
h3('4.5.2  室外训练场改造做法')
para('1.拆除原有破损塑化面层，清理场地基层，去除表面破损材料、杂物及松动土壤；')
para('2.基层平整、压实处理，确保基层坚实平整；')
para('3.重新铺设足球场人工草皮、塑胶跑道、拼装弹性运动地板，满足训练使用要求。')
h3('4.5.3  3号楼车位及相邻草坪硬化做法')
para('1.拆除草坪砖及绿植隔离带，场地平整、压实；')
para('2.铺设级配碎石垫层（500mm厚，第一层300mm厚级配碎石压实、第二层200mm厚5%水泥稳定碎石基层压实）；')
para('3.铺设沥青混凝土面层（50mm厚AC-20中粒式下面层+30mm厚AC-13细粒式上面层）；')
para('4.移建充电桩6个、路灯5个、监控2个，恢复场地功能。')

# 第五章
p = h1('第五章  安全设施设计')
add_bookmark(p, '_TOC_CS_5', 105)
h2('5.1  结构安全')
para('本项目为维修项目，不改变原建筑主体结构体系。空调主机基础加固及场地硬化施工严格按照现行规范要求实施，施工前对既有结构和场地进行现场查勘，明确受力状态，确保房屋整体安全。')
h2('5.2  施工安全')
para('1.施工前编制专项施工方案，落实安全生产责任制，对作业人员进行安全技术交底；')
para('2.空调主机吊装作业设置警戒区域，专人指挥，严禁无关人员进入吊装区域；')
para('3.塑胶场地施工及沥青摊铺作业做好防烫、防滑、防火措施，作业人员佩戴相应防护用品；')
para('4.施工现场用电由专业电工负责，严格执行临时用电安全规范，做到"一机一闸一漏一箱"；')
para('5.施工机械进场前检验合格，操作人员持证上岗；建筑垃圾吊运采用专用容器，严禁高空抛掷。')
h2('5.3  消防安全')
para('1.动火作业严格执行动火审批制度，配备灭火器材，作业后落实看火措施；')
para('2.施工期间保持疏散通道畅通，不占用、不堵塞消防设施，不改变原建筑消防分区和疏散体系；')
para('3.空调冷媒等材料按规范储存、使用，远离火源。')
h2('5.4  电气及给排水安全')
para('1.施工前办理停电手续，施工过程中不得损坏既有电气线路及设备，如发现损坏及时修复；')
para('2.临时用电线路敷设规范，配电箱设置漏电保护装置；')
para('3.空调系统安装调试由专业技术人员负责，确保电气连接安全可靠、系统运行正常。')

# 第六章
p = h1('第六章  节能设计')
add_bookmark(p, '_TOC_CS_6', 106)
h2('6.1  节能原则')
para('本项目为维修项目，在满足使用功能和现行标准的前提下，通过更新高效节能的多联机空调设备、选用节能环保材料等措施，实现节能降耗目标。')
h2('6.2  节能措施')
para('1.新装多联机空调采用直流变频技术，能效等级高，运行节能，较原老化设备可有效降低运行能耗；')
para('2.空调系统合理分区控制，按实际需求调节制冷制热，避免能源浪费；')
para('3.选用符合国家节能环保标准的建材，塑胶场地材料、沥青混凝土等满足环保要求；')
para('4.施工期间合理安排作业时段，减少能源消耗。')

# 第七章
p = h1('第七章  环境保护及影响评价')
add_bookmark(p, '_TOC_CS_7', 107)
h2('7.1  环境影响分析')
para('本项目施工期主要环境影响包括：拆除及场地作业产生的噪声、粉尘，建筑垃圾及废弃材料（废旧空调设备、破损塑胶面层、草坪砖、土方等）的运输处置，少量施工废水等。项目位于单位-LZ院内，周边为办公区域，施工期环境影响主要集中在施工阶段，采取有效措施后可控制在可接受范围内。')
h2('7.2  污染防治措施')
para('1.噪声控制：合理安排施工作业时间，避开机关办公时段；选用低噪声设备，高噪声作业采取隔声降噪措施；')
para('2.粉尘控制：施工区域设置围挡，拆除作业采用湿法作业，材料堆放覆盖，运输车辆密闭，出入口设置冲洗设施；')
para('3.废水控制：施工废水经沉淀处理后排放，不得随意排入市政管网；')
para('4.建筑垃圾处置：拆除的废旧空调设备、破损塑胶面层、草坪砖、土方等建筑垃圾分类收集、及时清运，按照《四川省住房和城乡建设厅关于规范建设工程建筑垃圾和拆除垃圾外运及处置费计价的通知》（川建标函〔2024〕3160号）等规定进行外运及处置，保持施工区域整洁，减少对周边环境及机关正常办公的影响；')
para('5.材料环保：选用符合国家环保标准的绿色建材，施工完成后对室内环境进行检测，确保空气质量达标。')

# 第八章
p = h1('第八章  项目实施组织管理')
add_bookmark(p, '_TOC_CS_8', 108)
h2('8.1  组织机构')
para('项目由单位-LZ组织实施，成立项目建设管理机构，明确专人负责项目建设管理，落实项目法人责任制，对项目建设的进度、质量、投资、安全实行全过程管理。')
h2('8.2  进度安排')
para('项目计划工期180日历天。拟按以下阶段组织实施：施工准备（含招标、合同签订、图纸会审）约30日历天；主体维修施工（空调系统更新、训练场改造、车位及草坪硬化）约120日历天；竣工验收及结算约30日历天。施工期间结合机关办公实际，合理安排作业时段，确保按期完工。')
h2('8.3  招标方案')
para('按照《中华人民共和国招标投标法》及四川省有关规定，本项目施工拟采用公开招标方式，委托具有相应资质的招标代理机构组织招标，招标范围为全部工程施工内容，招标组织形式为委托招标。设计、监理等服务按国家有关规定确定。招标方案如下表：')
add_table(
    ['招标内容', '招标范围', '招标组织形式', '招标方式', '估算金额（万元）'],
    [
        ['建筑工程施工', '全部招标', '委托招标', '公开招标', '273.71'],
        ['工程监理', '全部招标', '委托招标', '邀请招标或竞争性方式', '7.68'],
        ['设计', '全部招标', '委托招标', '竞争性方式', '10.14'],
    ],
    widths=[3.8, 2.4, 2.4, 3.0, 3.8],
)
para('', indent=False)
h2('8.4  合同管理')
para('严格执行合同管理制度，依法签订设计、施工、监理等合同，明确各方权利义务、质量标准、工期要求和违约责任；加强合同履约管理，规范工程变更签证程序。')
h2('8.5  质量与投资控制')
para('1.建立工程质量保证体系，实行工程监理制，关键工序、隐蔽工程（空调系统安装调试、场地基层压实等）实行旁站和验收制度，确保工程质量符合设计和规范要求；')
para('2.严格按照批准的投资概算控制工程投资，加强工程计量与支付管理，防止超概算；')
para('3.工程竣工后及时组织竣工验收，办理竣工结算和财务决算，接受审计监督。')

# 第九章
p = h1('第九章  投资概算与资金筹措')
add_bookmark(p, '_TOC_CS_9', 109)
h2('9.1  概算编制依据')
para('1.项目-LZ办公区综合维修改造项目设计图纸；')
para('2.《建设工程工程量清单计价规范》GB 50500-2013及《房屋建筑与装饰工程工程量计算规范》GB 50854-2013等现行工程量计算规范；')
para('3.《四川省建设工程工程量清单计价定额》（2020）及相关配套和解释文件；')
para('4.《四川省住房和城乡建设厅关于重新调整<建筑业营业税改征增值税四川省建设工程计价依据调整办法>的通知》（川建造价发〔2019〕181号）；')
para('5.人工费调整执行川建价发〔2025〕41号《关于批准2026年上半年各市（州）2020年〈四川省建设工程工程量清单计价定额〉人工费调整幅度及计日工人工单价的批复》；')
para('6.安全文明施工费执行《四川省建设工程安全文明施工费计价管理办法》（川建行规〔2024〕15号），其他总价措施取费标准执行川建标函〔2024〕3159号；')
para('7.材料价格采用2026年第6期《建设工程价格信息》成都市市区价格，无对应价格的采用市场询价；')
para('8.规费费率按照I档计取，销项增值税税率按9%计取，附加税税率按0.313%计取；')
para('9.《基本建设项目建设成本管理规定》（财建〔2016〕504号）、《建设工程监理与相关服务收费管理规定》（发改价格〔2007〕670号）、《工程勘察设计收费管理规定》（计价格〔2002〕10号）、《国家发展改革委关于降低部分建设项目收费标准规范收费行为等有关问题的通知》（川发改价格〔2011〕323号）、《招标代理服务收费管理暂行办法》（计价格〔2002〕1980号）、《工程造价咨询服务收费标准》（川价发〔2008〕141号）等有关取费规定。')
h2('9.2  概算范围')
para('本项目投资概算范围为经批准维修内容所包含的工程费用、工程建设其他费用和预备费，即：2号楼多联机空调系统更新、室外训练场改造、3号楼车位及相邻草坪硬化等全部维修内容。')
h2('9.3  项目投资概算')
para('项目概算总投资318.12万元，其中：第一部分建筑安装工程费用273.71万元，第二部分工程建设其他费用29.26万元，第三部分预备费15.15万元（基本预备费，按第一、二部分费用之和的5%计取）。投资概算汇总如下表：')
add_table(
    ['序号', '费用名称', '概算金额（万元）', '计算说明'],
    [
        ['一', '第一部分  建筑安装工程费用', '273.71', '其中装饰工程260.67万元、零星工程费13.03万元（按一类费用5%计取）'],
        ['二', '第二部分  工程建设其他费用', '29.26', ''],
        ['', '其中：建设单位管理费', '5.47', '参照财建〔2016〕504号'],
        ['', '工程监理费', '7.68', '参照发改价格〔2007〕670号'],
        ['', '设计费', '10.14', '参照计价格〔2002〕10号'],
        ['', '设计文件审查费', '0.55', '参照川发改价格〔2011〕323号'],
        ['', '工程招标代理费', '2.74', '参照计价格〔2002〕1980号'],
        ['', '招标工程量清单及控制价编制费', '1.33', '参照川价发〔2008〕141号'],
        ['', '竣工结算审核费', '1.37', '参照川价发〔2008〕141号'],
        ['三', '第三部分  预备费', '15.15', '按第一、二部分费用之和的5%'],
        ['四', '概算总投资', '318.12', ''],
    ],
    widths=[1.2, 6.2, 3.0, 5.0],
)
para('', indent=False)
para('建筑安装工程费用中，各分部分项工程费用构成如下表：')
add_table(
    ['序号', '分部工程名称', '金额（万元）', '主要内容'],
    [
        ['1', '2号楼多联机空调系统更新', '86.21', '直流变频室外机12台、线控器、连接铜管及风口、原主机拆除、室内机清洗等'],
        ['2', '室外训练场改造', '140.70', '足球场拆除、人工草皮、塑胶跑道、拼装弹性运动地板等'],
        ['3', '3号楼车位及相邻草坪硬化', '33.75', '草坪砖及隔离带拆除、级配碎石垫层、沥青混凝土面层、充电桩路灯监控移建等'],
        ['', '装饰工程小计', '260.66', ''],
        ['4', '零星工程费', '13.03', '按装饰工程一类费用的5%计取'],
        ['', '建筑安装工程费用合计', '273.71', ''],
    ],
    widths=[1.2, 5.0, 3.2, 6.0],
)
para('', indent=False)
para('注：上表尾数按四舍五入保留两位小数；工程垃圾、拆除垃圾外运及处置费按川建标函〔2024〕3160号规定由投标人自主报价、包干使用，结算时不作调整。')
h2('9.4  投资控制')
para('本项目概算总投资318.12万元，项目在实施过程中将严格执行批准的投资概算，加强投资控制，确保不突破概算。')
h2('9.5  资金筹措')
para('项目资金来源为省级财政资金，由项目单位按程序申请拨付。项目概算总投资318.12万元，资金需求纳入年度预算安排，确保项目顺利实施。')

# 第十章
p = h1('第十章  结论与建议')
add_bookmark(p, '_TOC_CS_10', 110)
h2('10.1  结论')
para('1.项目-LZ办公区综合维修改造项目（办公用房维修项目）维修内容符合《党政机关办公用房建设标准》（发改投资〔2014〕2674号）及《四川省省级机关办公用房维修管理实施办法（2023年修订版）》有关规定。')
para('2.项目针对2号楼空调系统老化失效、室外训练场塑胶铺装损坏、3号楼车位及草坪区域闲置等突出问题实施维修，技术可行、经济合理，实施必要。')
para('3.项目概算总投资318.12万元，资金来源为省级财政资金，资金有保障。')
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
