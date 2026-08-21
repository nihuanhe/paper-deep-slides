# paper-deep-slides（论文精读幻灯片生成器）

> Pi / Claude 技能（Skill）：把生物医学论文生成为**组会精讲用的学术幻灯片 PPTX**。
> 输入论文 PDF + 用户自总结的精读稿（含图）+ 可选补充截图，输出结构固定、版式统一的 16:9 汇报文稿。

English version: [README-EN.md](README-EN.md)

---

## 功能特性

- **按原文 Results 小标题组织**：每个小标题一个分隔页（中译标题 + 英文原题 + Fig 范围 + 本节问题 + 实验导读）
- **实验页全篇连续编号**（实验 1…N，跨章节不重置），每页固定四段：为什么做 / 做了什么 / 结果分析 / 结果意义
- **趋势对照式表达（强制）**：结果只写"实验组 vs 对照组"的变化方向（↑↓/→ 箭头；红 = 处理组变化与关键结论，蓝 = 对照/阴性/无差异），**不写具体数值、倍数和 P 值**
- **图片自适应最大化**：宽幅图（宽高比 ≥ 2.2）全宽展示、结果置于图下；中高图满高展示、结果置于图右，无需手调
- **设计系统统一**：米白底 + 深海军蓝 + 暗砖红，微软雅黑 + Arial
- **PowerPoint COM 渲染自检（S5）**：逐页导出 PNG 检查文字溢出、孤行、图文遮挡

## 目录结构

```
paper-deep-slides/
├── README.md                 # 语言入口索引
├── README-ZH.md              # 中文说明（本文件）
├── README-EN.md              # English documentation
├── SKILL.md                  # 技能主文档（输入/输出/流程 S1–S5）
├── assets/
│   └── deck_kit.py           # 设计系统工具库（颜色/字体/组件：标题栏、分隔页、实验页等）
└── references/
    ├── content-rules.md      # 内容写作规范（趋势对照式表达、编号规则、分隔页规则）
    └── slides-example.py     # 内容组织骨架示例（复制到工作目录后按论文改写）
```

## 依赖与配置

### 1. 系统要求

| 项 | 要求 | 说明 |
|---|---|---|
| 操作系统 | Windows 10 / 11 | S5 渲染自检依赖 PowerPoint COM；其余步骤跨平台可用 |
| Microsoft PowerPoint | 2016 或更新 | **仅用于 S5 渲染自检**（逐页导出 PNG）。无 PowerPoint 时可跳过 S5 或改用 LibreOffice 渲染 |
| Python | 3.9+ | 已测试 3.11–3.13 |

### 2. Python 依赖（pip 包）

```bash
pip install python-pptx pillow pymupdf python-docx
```

| 包 | 用途 |
|---|---|
| `python-pptx` | 生成 / 读取 .pptx 文件（核心） |
| `pillow` | 读取图片尺寸、实现"图片自适应最大化"布局 |
| `pymupdf` (fitz) | 从论文 PDF 提取全文，定位 Results 小标题与图版范围 |
| `python-docx` | 解析精读稿 DOCX，提取内嵌图片并语义化命名 |

### 3. 环境验证

```bash
python -c "import pptx, PIL, fitz, docx; print('deps ok')"
```

预期输出 `deps ok`。若报错，按上方 pip 命令补装对应包。

### 4. 字体

- 中文字体使用**微软雅黑**（Microsoft YaHei），Windows 自带，无需额外安装
- 西文使用 Arial，Windows 自带
- 若在 macOS/Linux 使用，请先安装 `msyh`（微软雅黑）字体，否则中文可能回退为默认字体

### 5. 模型建议

- 建议在**强模型**（如 Claude Opus 级别）下运行本技能
- 弱模型容易生成平淡叙事、忽略"趋势对照式表达"规则

## 使用方法

在 Pi / Claude 中加载技能并调用：

```
/skill:paper-deep-slides
```

然后提供：

1. **论文 PDF 路径**（必需）：用于提取 Results 小标题、图版范围、术语、图注
2. **精读稿路径**（必需）：DOCX 或 Markdown，通常内嵌已裁好的 panel 图，是内容主源
3. **补充图片路径**（可选）：未插入精读稿的散图

技能按五步执行：

```
S1 建项目目录、备素材   → 提取 DOCX 内嵌图、PDF 转文本、定位 Results 小标题
S2 规划结构（先确认）    → 分隔页/实验页映射表、证据链环节、实验连续编号
S3 内容改写（严格规则）  → 四段式 + 趋势对照式表达，bullet ≤ 30 字
S4 生成幻灯片           → 复制 deck_kit.py，编写 slides_a/b + build_pptx.py，运行生成
S5 渲染自检（必做）     → PowerPoint COM 逐页导出 PNG，人工逐页检查
```

## 示例成品

- 微信公众号「小云科研」图文《应激经脑-骨髓诱导巨噬细胞促胶质瘤生长-1》

## 常见问题（FAQ）

**Q1：pip 安装 pymupdf 失败？**
确保 Python 版本 ≥ 3.9；必要时使用 `python -m pip install pymupdf`。

**Q2：没有 PowerPoint 怎么办？**
S5 渲染自检可跳过（不导出 PNG），或安装 LibreOffice 后用其命令行导出检查。

**Q3：生成的中文变成方块/乱码？**
检查系统是否安装微软雅黑；确认在 python-pptx 中字体已设为"微软雅黑"（deck_kit 已默认处理）。

**Q4：图片被压得太小？**
deck_kit 的 `experiment_slide` 已做自适应最大化；若仍偏小，检查原始截图宽高比（截图应尽量贴合 panel 区域，避免大片留白）。
