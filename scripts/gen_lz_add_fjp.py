# -*- coding: utf-8 -*-
"""
项目-LZ项目：
1) 初步设计文档 4.4 主要材料或设备表 —— 新增 2 套风机盘管(FGR5Pd/FaNh-N1)及配套室外机(FGR5Pd/Na-N1)，并重新编号；
2) 项目申报审批模板 —— 「三、项目内容及规模」分项1 补充风机盘管及配套室外机描述。
"""
import copy
from docx import Document
from docx.oxml.ns import qn

DOC = r"<桌面>/改造项目PDF图纸/改造项目PDF图纸/项目-LZ办公区综合维修改造项目/项目-LZ办公区综合维修改造项目初步设计（代项目建议书、可行性研究报告）.docx"
TMPL = r"<桌面>/改造项目PDF图纸/改造项目PDF图纸/项目-LZ办公区综合维修改造项目/项目申报审批模板.docx"

FONT = "仿宋_GB2312"


def fill_tr(tr, values):
    """用 XML 填充新插入行：保留第一个 run 的字体格式，只改文字。"""
    tcs = tr.findall(qn("w:tc"))
    for tc, val in zip(tcs, values):
        p = tc.find(qn("w:p"))
        runs = p.findall(qn("w:r"))
        if not runs:
            r = p.makeelement(qn("w:r"), {})
            t = p.makeelement(qn("w:t"), {})
            t.text = val
            t.set(qn("xml:space"), "preserve")
            r.append(t)
            p.append(r)
            continue
        r0 = runs[0]
        for r in runs[1:]:
            p.remove(r)
        for t in r0.findall(qn("w:t")):
            r0.remove(t)
        t = r0.makeelement(qn("w:t"), {})
        t.text = val
        t.set(qn("xml:space"), "preserve")
        r0.append(t)


def set_cell_text(cell, text):
    """python-docx 版：清空多余 run，第一个 run 承载新文字（保留字体）。"""
    para = cell.paragraphs[0]
    runs = para.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ""
    else:
        run = para.add_run(text)
        run.font.name = FONT
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)


def main():
    # ---------- 任务1a：插入两行 ----------
    doc = Document(DOC)
    tb = doc.tables[3]
    ref_tr = tb.rows[8]._tr  # 序号8「单/双层百叶风口」

    tr_fjp = copy.deepcopy(ref_tr)  # 风机盘管
    tr_wjj = copy.deepcopy(ref_tr)  # 配套室外机
    ref_tr.addnext(tr_wjj)
    ref_tr.addnext(tr_fjp)  # 顺序：ref_tr, tr_fjp, tr_wjj

    fill_tr(tr_fjp, ["9", "风机盘管", "FGR5Pd/FaNh-N1", "台", "2", "空调系统"])
    fill_tr(tr_wjj, ["10", "配套室外机", "FGR5Pd/Na-N1", "台", "2", "空调系统"])
    doc.save(DOC)
    print("已插入两行并保存")

    # ---------- 任务1b：重新打开，统一重排序号 ----------
    d2 = Document(DOC)
    tb2 = d2.tables[3]
    for i, row in enumerate(tb2.rows[1:], start=1):
        set_cell_text(row.cells[0], str(i))
    d2.save(DOC)
    print("序号已重排")

    # ---------- 验证 ----------
    d3 = Document(DOC)
    tb3 = d3.tables[3]
    print("=== 更新后 4.4 表（共 %d 行）===" % len(tb3.rows))
    for r in tb3.rows:
        print(" | ".join(c.text.strip() for c in r.cells))

    # ---------- 任务2：申报审批模板 ----------
    tmpl = Document(TMPL)
    old = "原主机拆除吊装、室内机清洗检测）"
    new = "原主机拆除吊装、室内机清洗检测；1楼卫生间新增风机盘管FGR5Pd/FaNh-N1及配套室外机FGR5Pd/Na-N1共2套）"
    done = False
    for para in tmpl.paragraphs:
        if old in para.text:
            runs = para.runs
            if runs:
                runs[0].text = para.text.replace(old, new)
                for r in runs[1:]:
                    r.text = ""
            done = True
            break
    tmpl.save(TMPL)
    print("\n申报审批模板更新:", "OK" if done else "MISS")


if __name__ == "__main__":
    main()
