---
name: paper-deep-slides
description: 把生物医学论文（PDF 原文 + 用户自总结的 DOCX 精读稿 + 可选的补充图片）生成为组会精讲用的学术幻灯片 PPTX。用户提供论文 PDF 和带图的 DOCX 总结、说"做组会幻灯片/精读 PPT/把 DOCX 转成幻灯片"时使用。结构按原文 Results 小标题组织（分隔页），每个实验一页且全篇连续编号；结果用"实验组 vs 对照组"趋势表达（↑↓箭头、红蓝编码），不放具体数值和 P 值；图片自适应最大化展示。按 S1→S5 顺序执行。
---

# 论文精读幻灯片生成器（生物医学组会）

## 输入

- **论文 PDF**（必需）：提取 Results 小标题、图版范围（Fig. nX–Y）、术语、图注
- **DOCX 精读稿**（必需）：内容主源，通常内嵌已裁好的 panel 图
- **补充图片**（可选）：未插入 DOCX 的散图，用户单独给出路径

## 输出

单份 PPTX（16:9），结构固定：标题页 → 背景与科学问题 → 证据链总览 →
**按原文 Results 小标题各一个分隔页** → 每个实验一页（**全篇连续编号：实验 1…N，跨章节不重置**）→
全文总结 → 机制模型 → 讨论与展望。

## S1 建项目目录、备素材

在 DOCX 同级建工作目录（如 `<论文名>-slides\`），内含 `images\`：

1. **提取 DOCX 内嵌图**：
   ```bash
   cd <工作目录> && unzip -o -j "<精读稿.docx>" "word/media/*" -d images/
   ```
   DOCX 中图片出现顺序通常与精读顺序一致。用 python-docx 遍历段落，按图注/上下文把
   `image1.png` 等重命名为 `fig1_ABC.jpg` 这类语义名（便于 slides 代码引用）。
2. **复制补充图片**进 `images\`，同样语义化命名。
3. **PDF 转文本**：`python -c "import fitz; ..."`（pymupdf）提取全文，
   定位 `Results` 下所有小标题（#### 级）及每个小标题覆盖的图版范围；
   同时记录每个 panel 的图注（写 caption 用）。

## S2 规划结构（先给用户确认再动手）

列出映射表请用户确认：

- 每个 **Results 小标题** → 一个分隔页（中译标题 + 英文原题 + Fig 范围 + 本节问题 + 实验导读）
- 小标题下的**每个实验** → 一个实验页；实验编号 = 全篇连续序号（Fig.2 的第一个实验接着 Fig.1 的最后一个往下排）
- 归纳论文的**证据链环节**（3–6 步，如 磁性来源→细胞身份→功能验证→神经联系→行为验证），
  给每个 Results 小标题标注所属环节索引（分隔页底部进度条用）

## S3 内容改写（严格遵守）

读 [references/content-rules.md](references/content-rules.md) 并逐条执行。要点：

- 每实验页四段：**为什么做 / 做了什么 / 结果分析 / 结果意义**
- 结果一律写成 **实验组 vs 对照组的变化趋势**（↑↓/→ 箭头；红=处理组变化与关键结论，
  蓝=对照/阴性/无差异），**不写具体数值和 P 值**
- bullet ≤ 30 字、每段 ≤ 3 条；结果意义只写一句话
- 分隔页**不放任何结果图**（防止听众误以为重复）

## S4 生成幻灯片

1. 复制 `assets/deck_kit.py` 到工作目录（设计系统：米白底+深海军蓝+暗砖红，微软雅黑+Arial）。
2. 参照 [references/slides-example.py](references/slides-example.py) 编写
   `slides_a.py` / `slides_b.py` / `build_pptx.py`：
   - `title_bar` / `takeaway_bar` / `divider_slide` / `experiment_slide` 组件直接调用
   - `experiment_slide` 内置**自适应图片最大化**（宽幅图 r≥2.2 全宽、结果在图下；
     中高图占满正文高、结果在图右），无需手调
   - `divider_slide(..., chain=[环节...], active=(索引,))` 画进度条
3. `python build_pptx.py` 生成。

## S5 渲染自检（必须执行）

用 PowerPoint COM 把每页导出 PNG 后**逐页用 read 查看**：

```powershell
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('<绝对路径.pptx>', $true, $false, $false)
for ($i=1; $i -le $pres.Slides.Count; $i++) {
  $pres.Slides.Item($i).Export("<工作目录>\render\slide$i.png", "PNG", 1600, 900) }
$pres.Close(); $ppt.Quit()
```

检查：文字溢出/孤行（单字独行）、图文遮挡、图片是否被压得过小、长 bullet 换行难看。
发现问题改 slides_*.py 重生成重渲染，直到全部通过。

## 备注

- 环境要求：Windows + PowerPoint（自检渲染用）、`pip install python-pptx pillow pymupdf python-docx`
- 标题页汇报人/日期用占位符，提醒用户自行填写
- 用户的原始 DOCX/PDF 初稿等文件一律不修改
