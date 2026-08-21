# paper-deep-slides

Pi / Claude 技能（Skill）：把生物医学论文生成为组会精讲用的学术幻灯片 PPTX。

## 功能

输入 **论文 PDF（原文）+ 用户自总结的精读稿（DOCX/MD，含图）+ 可选补充截图**，输出一份结构固定的 16:9 学术汇报 PPTX：

- **按原文 Results 小标题组织**：每个小标题一个分隔页（中译标题 + 英文原题 + Fig 范围 + 本节问题 + 实验导读）
- **实验页全篇连续编号**（实验 1…N，跨章节不重置），每页四段：为什么做 / 做了什么 / 结果分析 / 结果意义
- **趋势对照式表达**：结果只写"实验组 vs 对照组"变化方向（↑↓/→ 箭头，红=处理组/关键结论，蓝=对照/阴性），**不放具体数值与 P 值**
- **图片自适应最大化**：宽幅图（r≥2.2）全宽展示、结果在图下；中高图满高展示、结果在图右
- **PowerPoint COM 渲染自检**：逐页导出 PNG 检查文字溢出、图文遮挡

## 目录结构

```
paper-deep-slides/
├── SKILL.md                     # 技能主文档（输入/输出/流程 S1–S5）
├── assets/
│   └── deck_kit.py              # 设计系统工具库（米白底+深海军蓝+暗砖红，微软雅黑+Arial）
└── references/
    ├── content-rules.md         # 内容写作规范（趋势对照式表达、编号、分隔页规则）
    └── slides-example.py        # 内容组织骨架示例
```

## 使用方法

在 Pi / Claude 中加载技能并调用：

```
/skill:paper-deep-slides
```

然后提供：
1. 论文 PDF 路径
2. 精读稿路径（DOCX 或 Markdown，通常内嵌已裁好的 panel 图）
3. （可选）补充图片路径

技能按 **S1 建目录备素材 → S2 规划结构（确认后动手）→ S3 内容改写 → S4 生成幻灯片 → S5 渲染自检** 五步执行。

## 环境要求

- Windows + PowerPoint（S5 渲染自检用 COM）
- Python 包：`pip install python-pptx pillow pymupdf python-docx`
- 建议在强模型（如 Claude Opus 级）下运行，弱模型易生成平淡叙事

## 示例成品

- Science 2026《鸽子巨噬细胞导航》：`Science2026-鸽子-巨噬细胞-导航-final.pptx`（24 页）
- Cancer Cell 2026《慢性应激经脑-骨髓交互促进胶质瘤生长》：`Yang2026-SAMs-组会精读.pptx`（41 页）
