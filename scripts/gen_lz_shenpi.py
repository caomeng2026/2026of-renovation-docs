# -*- coding: utf-8 -*-
"""
生成「项目-LZ办公区综合维修改造项目」项目申报审批模板（项目审批申请模版）
以「单位-ZX主楼综合维修改造项目」已完成模板为基底，精确替换文字，保持既有格式（宋体）。
"""
import shutil
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

SRC = r"<桌面>/改造项目PDF图纸/改造项目PDF图纸/单位-ZX主楼综合维修改造项目/项目申报审批模板.docx"
DST = r"<桌面>/改造项目PDF图纸/改造项目PDF图纸/项目-LZ办公区综合维修改造项目/项目申报审批模板.docx"

# 目标替换内容（按原文唯一关键词定位）
REPLACEMENTS = [
    ("关于申请审批", "关于申请审批项目-LZ办公区综合维修改造项目初步设计（代项目建议书、可行性研究报告）和投资概算的函"),
    ("已列入年度维修安排计划", "我厅（局）项目-LZ办公区综合维修改造项目已列入年度维修安排计划，按照《四川省省级机关办公用房维修管理实施办法（2023年修订版）》，已完成前期资料准备，现提出项目初步设计（代项目建议书、可行性研究报告）和投资概算审批申请。"),
    ("单位-ZX主楼综合维修改造项目。", "项目-LZ办公区综合维修改造项目。"),
    ("地址-ZX", "地址-LZ。"),
    ("主楼房门/防火门更换", "1、2号楼多联机空调系统更新（采购并安装全新直流变频多联机空调室外机5种型号共12台、线控器187个、连接铜管φ6.4—28.6mm及单/双层百叶风口，原主机拆除吊装、室内机清洗检测）"),
    ("议政厅屋面防水改造", "2、室外训练场改造（塑化面层全面重新铺装约3000㎡：足球场人工草皮约1200㎡、塑胶跑道约2500㎡、拼装弹性运动地板约600㎡，原破损面层拆除、基层平整压实）"),
    ("负二楼出入口防水维修改造", "3、3号楼车位及相邻草坪硬化（现有26个地面机动车位及相邻草坪统一硬化：沥青混凝土面层644㎡、级配碎石垫层386.4m³、水泥稳定碎石基层193.2m³；移建充电桩6个、路灯5个、监控2个）。"),
    ("投资概算154.04万元", "投资概算318.12万元，未超计划投资控制限额（321万）。"),
    ("XX单位（中心）", "单位-LZ。"),
]


def replace_para_text(para, new_text):
    """清空段落所有 run，用第一个 run 承载新文本，字体统一宋体、字号/加粗沿用原段落。"""
    runs = para.runs
    if not runs:
        run = para.add_run(new_text)
        run.font.name = "宋体"
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
        return
    first = runs[0]
    size = first.font.size.pt if first.font.size else None
    bold = first.font.bold
    first.text = new_text
    first.font.name = "宋体"
    if size:
        first.font.size = Pt(size)
    if bold is not None:
        first.font.bold = bold
    rPr = first._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), "宋体")
    rFonts.set(qn("w:ascii"), "宋体")
    rFonts.set(qn("w:hAnsi"), "宋体")
    for r in runs[1:]:
        r.text = ""


def main():
    shutil.copy(SRC, DST)
    doc = Document(DST)
    hit = {k: False for k, _ in REPLACEMENTS}

    for para in doc.paragraphs:
        full = para.text
        for key, new_text in REPLACEMENTS:
            if key in full:
                replace_para_text(para, new_text)
                hit[key] = True
                break

    doc.save(DST)

    print("=== 替换结果 ===")
    for k, ok in hit.items():
        print(f"[{'OK' if ok else 'MISS'}] {k}")
    print("\n保存至:", DST)


if __name__ == "__main__":
    main()
