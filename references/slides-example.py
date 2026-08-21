# -*- coding: utf-8 -*-
"""内容组织骨架示例：复制到工作目录后按论文实际内容改写。
deck_kit.py 由 skill 的 assets/ 复制而来，与本文件同目录。

run 样式：n=正文 b=加粗 r=红 rb=红粗(处理组/关键结论) bl=蓝粗(对照/阴性) g=灰
每条 bullet 是 run 列表：[(文本, 样式), ...]，可在一句内混色。
"""
import os
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from deck_kit import (NAVY, NAVY_DEEP, RED, BG, INK, GRAY, LIGHT_RED, LIGHT_BLU,
                      BORDER, WHITE, SOFT_WHT, SLIDE_W, new_slide, add_text,
                      add_rect, add_pic_fit, page_number, title_bar,
                      takeaway_bar, experiment_slide, divider_slide)

# 证据链环节（按论文归纳，3–6 步），分隔页进度条 & 总览页共用
CHAIN = ["① 环节A", "② 环节B", "③ 环节C", "④ 环节D"]


def build(prs, IMG):
    # ---------- S1 标题页（仿照：期刊tag/中文大标题/英文题/作者/汇报人占位/题图/底栏链条）----------
    s = new_slide(prs)
    add_rect(s, 0, 0, SLIDE_W, 0.14, fill=NAVY)
    add_text(s, 0.9, 1.05, 6.6, 0.4,
             [dict(runs=[("Science · 2026 · 组会论文精读", "r")], size=13, bold=True)])
    add_text(s, 0.9, 1.55, 6.6, 1.9,
             [dict(runs=[("中文标题第一行", "n")], size=33, color=NAVY, bold=True, space_after=2),
              dict(runs=[("中文标题第二行", "n")], size=33, color=NAVY, bold=True)])
    add_text(s, 0.9, 3.35, 6.6, 1.0,
             [dict(runs=[("English paper title", "n")], size=13.5, color=GRAY, italic=True, line=1.2)])
    add_text(s, 0.9, 4.35, 6.6, 0.4, [dict(runs=[("作者  |  期刊 (年份)", "g")], size=11.5)])
    add_text(s, 0.9, 4.85, 6.6, 0.4, [dict(runs=[("汇报人：（姓名）  ·  20XX-XX-XX", "g")], size=11.5)])
    # add_pic_fit(s, os.path.join(IMG, "title.jpg"), 7.6, 1.35, 5.0, 4.3, border=True)
    add_rect(s, 0, 6.55, SLIDE_W, 0.95, fill=NAVY)
    add_text(s, 0, 6.55, SLIDE_W, 0.95,
             [dict(runs=[("  →  ".join(CHAIN), "n")], size=13, color=WHITE, bold=True)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # ---------- S4 Results 小标题分隔页（不放结果图！chain+active 画进度条）----------
    divider_slide(prs, 4, "RESULTS · 1",
                  "小标题中译（大字）",
                  "Original English subheading  ·  Fig. 1A–D",
                  "本节核心问题？",
                  "两个实验：方法A → 方法B",
                  active=(0,), chain=CHAIN)

    # ---------- 实验页（编号全篇连续；图自适应最大化）----------
    experiment_slide(prs, IMG, 5, "证据链 ① 环节A", "实验 1 · 实验名", "（Fig. 1A–C）",
        why=[[("动机一句话，", "n"), ("关键术语", "b"), ("……", "n")]],
        how=[[("样本/处理 → ", "n"), ("检测手段", "b")],
             [("对照设置", "n")]],
        results=[[("对照组 → ", "n"), ("处理组趋势", "rb"), (" ↑↓", "rb")],
                 [("另一指标：", "n"), ("无显著差异", "bl")]],
        takeaway=[("一句话结论，", "n"), ("关键结论", "rb")],
        img="fig1_ABC.jpg",
        caption="Fig. 1A–C：各 panel 一句话要点（保留 panel 字母）")

    # ---------- 后续：更多 divider + experiment；总结/机制/讨论页仿此用 add_text/add_rect 搭建 ----------
