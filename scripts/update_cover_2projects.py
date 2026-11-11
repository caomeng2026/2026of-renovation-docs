# -*- coding: utf-8 -*-
"""按新封面格式标准修复两份旧文档封面：
1. 单位-ZY：初步设计36->26pt、代..26->22pt、信息行加粗、补建设地点+设计单位
2. 单位-ZX主楼：信息行加粗、补设计单位
新格式标准：
- 主标题=项目全称 26pt 加粗居中（一行）
- "初步设计" 26pt 加粗居中
- "（代项目建议书、可行性研究报告）" 22pt 加粗居中
- 信息行 16pt 加粗：项目单位/建设地点/设计单位(留白16全角空格，两端对齐+首行缩进4字符)/编制时间
- 全文仿宋_GB2312
"""
import docx
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

FONT = '仿宋_GB2312'


def set_run(run, size, bold):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = FONT
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rPr.insert(0, rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rf.set(qn(a), FONT)


def make_info_para(template_p, text, justify=False, indent4=False):
    """以 template_p 的 pPr 为模板，生成一个信息行段落元素"""
    new_p = OxmlElement('w:p')
    ppr = copy.deepcopy(template_p._p.pPr)
    if ppr is not None:
        new_p.append(ppr)
    pPr = new_p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        new_p.insert(0, pPr)
    # 对齐
    jc = pPr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        pPr.append(jc)
    jc.set(qn('w:val'), 'both' if justify else 'center')
    # 首行缩进4字符（仅设计单位行）
    if indent4:
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            pPr.append(ind)
        ind.set(qn('w:firstLineChars'), '400')
        ind.set(qn('w:firstLine'), '1285')
    # run
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rf.set(qn(a), FONT)
    rPr.append(rf)
    b = OxmlElement('w:b')
    rPr.append(b)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '32')  # 16pt
    rPr.append(sz)
    szcs = OxmlElement('w:szCs')
    szcs.set(qn('w:val'), '32')
    rPr.append(szcs)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    new_p.append(r)
    return new_p


def locate_cover(d):
    """定位封面关键段落（前 35 段内）"""
    found = {'chubu': None, 'dai': None, 'danwei': None, 'didian': None, 'time': None}
    for p in d.paragraphs[:35]:
        t = p.text.strip()
        if not t:
            continue
        if t == '初步设计':
            found['chubu'] = p
        elif t.startswith('（代项目建议书'):
            found['dai'] = p
        elif t.startswith('项目单位'):
            found['danwei'] = p
        elif t.startswith('建设地点'):
            found['didian'] = p
        elif t.startswith('编制时间'):
            found['time'] = p
    return found


def fix_cover(path, need_didian_text=None):
    d = docx.Document(path)
    f = locate_cover(d)
    # 1) 字号修正
    if f['chubu'] is not None:
        for r in f['chubu'].runs:
            set_run(r, 26, True)
    if f['dai'] is not None:
        for r in f['dai'].runs:
            set_run(r, 22, True)
    # 2) 信息行加粗（项目单位/建设地点/编制时间）
    for key in ('danwei', 'didian', 'time'):
        if f[key] is not None:
            for r in f[key].runs:
                set_run(r, 16, True)
    # 3) 插入缺失的信息行（在"编制时间"之前）
    time_p = f['time']
    if time_p is not None:
        # 缺建设地点则补（居中）
        if need_didian_text and f['didian'] is None:
            dd = make_info_para(time_p, need_didian_text, justify=False, indent4=False)
            time_p._p.addprevious(dd)
        # 缺设计单位则补（两端对齐 + 首行缩进4字符 + 16全角空格）
        design = make_info_para(time_p, '设计单位：' + '\u3000' * 16, justify=True, indent4=True)
        time_p._p.addprevious(design)
    d.save(path)
    return d


files = {
    '单位-ZY': {
        'path': '<桌面>/改造项目PDF图纸/改造项目PDF图纸/单位-ZY科研办公大楼综合维修改造项目/提交盖章资料/单位-ZY科研办公大楼综合维修改造项目初步设计（代项目建议书、可行性研究报告）.docx',
        'didian': '建设地点：成都市武侯区人民南路四段51号',
    },
    '单位-ZX主楼': {
        'path': '<桌面>/改造项目PDF图纸/改造项目PDF图纸/单位-ZX主楼综合维修改造项目/提交盖章资料/单位-ZX主楼综合维修改造项目初步设计（代项目建议书、可行性研究报告）.docx',
        'didian': None,  # 已有建设地点行，无需补
    },
}

for name, cfg in files.items():
    fix_cover(cfg['path'], cfg['didian'])
    print(f'[{name}] 封面已修复')

print('全部完成')
