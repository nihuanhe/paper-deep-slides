# paper-deep-slides

> A Pi / Claude skill that turns a biomedical paper into a lab-meeting presentation (PPTX).

| Language | 语言 | File |
|---|---|---|
| 🇨🇳 简体中文 | 中文说明（含配置与依赖） | [**README-ZH.md**](README-ZH.md) |
| 🇬🇧 English | English docs (with setup & dependencies) | [**README-EN.md**](README-EN.md) |

---

**Quick start — one line:**

```bash
pip install python-pptx pillow pymupdf python-docx
```

Input: paper PDF + deep-reading notes (DOCX/MD with figures) + optional screenshots.
Output: fixed-structure 16:9 academic PPTX — Results-subheading dividers, consecutively numbered experiment pages, trend-comparison results (↑↓, red/blue coding), adaptive image maximization, PowerPoint COM render self-check.

See **README-ZH.md** (中文) or **README-EN.md** (English) for full details.
