# paper-deep-slides — Paper Deep-Reading Slide Generator

> A Pi / Claude skill that turns a biomedical paper into a **lab-meeting presentation (PPTX)**.
> Input: paper PDF + your own deep-reading notes (with figures) + optional extra screenshots.
> Output: a fixed-structure, uniform-styled 16:9 academic slide deck.

中文版：[README-ZH.md](README-ZH.md)

---

## Features

- **Organized by the paper's Results subheadings**: one divider slide per subheading (Chinese title + original English title + figure range + key question + experiment guide)
- **Globally consecutive experiment numbering** (Experiment 1…N, never resets across sections); every experiment page has four blocks: *Why / What / Result analysis / Takeaway*
- **Trend-comparison expression (mandatory)**: results are written only as directional changes of *treatment vs control* (↑↓/→ arrows; red = treated group / key conclusion, blue = control / negative / no difference); **no numeric values, fold-changes, or p-values**
- **Adaptive image maximization**: wide images (aspect ratio ≥ 2.2) span full width with results below; medium/tall images fill the content height with results on the right — no manual tuning
- **Unified design system**: off-white background + deep navy + brick red; Microsoft YaHei + Arial
- **PowerPoint COM render self-check (S5)**: exports every slide to PNG to catch text overflow, orphan lines, and text/figure overlap

## Directory Structure

```
paper-deep-slides/
├── README.md                 # Language entry index
├── README-ZH.md              # 中文说明
├── README-EN.md              # English documentation (this file)
├── SKILL.md                  # Skill main doc (inputs / outputs / S1–S5 workflow)
├── assets/
│   └── deck_kit.py           # Design-system toolkit (colors/fonts/components: title bar, divider, experiment page…)
└── references/
    ├── content-rules.md      # Content writing rules (trend-comparison style, numbering, divider rules)
    └── slides-example.py     # Content skeleton example (copy into your working dir and adapt)
```

## Dependencies & Setup

### 1. System Requirements

| Item | Requirement | Notes |
|---|---|---|
| OS | Windows 10 / 11 | S5 render self-check uses PowerPoint COM; other steps are cross-platform |
| Microsoft PowerPoint | 2016 or newer | **Only needed for the S5 render self-check** (export PNGs). Skip S5 or use LibreOffice if unavailable |
| Python | 3.9+ | Tested on 3.11–3.13 |

### 2. Python Dependencies (pip)

```bash
pip install python-pptx pillow pymupdf python-docx
```

| Package | Purpose |
|---|---|
| `python-pptx` | Create / read .pptx files (core) |
| `pillow` | Read image dimensions; drives the adaptive image-maximization layout |
| `pymupdf` (fitz) | Extract full text from the paper PDF; locate Results subheadings and figure ranges |
| `python-docx` | Parse the deep-reading DOCX; extract embedded figures and rename them semantically |

### 3. Environment Verification

```bash
python -c "import pptx, PIL, fitz, docx; print('deps ok')"
```

Expected output: `deps ok`. If it fails, install the missing package with the pip command above.

### 4. Fonts

- Chinese text uses **Microsoft YaHei (微软雅黑)** — bundled with Windows, no extra install
- Latin text uses Arial — bundled with Windows
- On macOS/Linux, install `msyh` (Microsoft YaHei) first, otherwise Chinese may fall back to a default font

### 5. Model Recommendation

- Run this skill with a **strong model** (e.g., Claude Opus class)
- Weak models tend to produce flat narratives and may ignore the trend-comparison rule

## Usage

Load the skill in Pi / Claude and invoke:

```
/skill:paper-deep-slides
```

Then provide:

1. **Paper PDF path** (required): extract Results subheadings, figure ranges, terminology, captions
2. **Deep-reading notes path** (required): DOCX or Markdown, usually containing pre-cropped panel figures; the primary content source
3. **Extra figures path** (optional): loose screenshots not embedded in the notes

The skill runs in five steps:

```
S1 Prepare workspace & assets → extract DOCX images, PDF→text, locate Results subheadings
S2 Plan structure (confirm first) → divider/experiment mapping, evidence-chain steps, consecutive numbering
S3 Rewrite content (strict rules) → four blocks + trend-comparison style; bullets ≤ 30 chars
S4 Generate slides → copy deck_kit.py, write slides_a/b + build_pptx.py, run
S5 Render self-check (mandatory) → PowerPoint COM exports PNGs; review page by page
```

## Example Outputs

- Science 2026, "Pigeon macrophage navigation": `Science2026-鸽子-巨噬细胞-导航-final.pptx` (24 slides)
- Cancer Cell 2026, "Chronic stress → brain-bone marrow crosstalk in glioma": `Yang2026-SAMs-组会精读.pptx` (41 slides)

## FAQ

**Q1: `pip install pymupdf` fails?**
Make sure Python ≥ 3.9; try `python -m pip install pymupdf`.

**Q2: No PowerPoint available?**
Skip S5 (no PNG export), or install LibreOffice and render via its CLI for a visual check.

**Q3: Chinese text renders as boxes/mojibake?**
Check that Microsoft YaHei is installed; deck_kit already sets "微软雅黑" in python-pptx by default.

**Q4: Images come out too small?**
deck_kit's `experiment_slide` maximizes images automatically; if still small, check the source screenshots' aspect ratio (crop tightly to the panel area, avoid large white margins).
