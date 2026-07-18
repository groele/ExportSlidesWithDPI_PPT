<div align="center">

# ExportSlidesWithDPI_PPT

**面向论文图件与学术展示的 PowerPoint 高 DPI 幻灯片导出插件**  
*PowerPoint export utility for high-DPI images and PDFs, selected-page export, and publication/presentation-ready visual assets.*

![Type](https://img.shields.io/badge/type-PowerPoint%20Add--in-blue?style=flat-square)
![Domain](https://img.shields.io/badge/domain-slide%20export%20%2F%20figures-green?style=flat-square)
![Language](https://img.shields.io/badge/language-C%23-blueviolet?style=flat-square)
![Architecture](https://img.shields.io/badge/architecture-Office%20Interop-purple?style=flat-square)
![License](https://img.shields.io/badge/license-MIT-yellow?style=flat-square)

Part of **ResearchFlow Lab** — a local-first research productivity ecosystem for literature, manuscripts, data, and scientific visualization.

</div>

---

## 01. Overview

**ExportSlidesWithDPI_PPT** is a PowerPoint utility for exporting selected slides, custom page ranges, or complete presentations as high-DPI images. It is designed for academic figures, conference slides, graphical abstracts, mechanism diagrams, and manuscript-supporting visual material.

**ExportSlidesWithDPI_PPT** 是一个 PowerPoint 高 DPI 图片导出插件，支持当前页、指定页码范围和全部幻灯片导出，适合论文图件、学术展示、机制示意图和投稿图像材料准备。

---

## 02. Why this project exists

PowerPoint is widely used to assemble scientific diagrams and presentation figures, but its default export workflow is often insufficient for high-resolution academic use. Researchers frequently need exact page-range export, consistent DPI, and reproducible image formats for manuscripts, presentations, covers, posters, and supplementary figures.

核心目标：

- Export PowerPoint slides with user-defined DPI.
- Support current slide, custom page ranges, and full-presentation export.
- Support common image formats such as PNG, JPG, and TIFF.
- Provide a more controlled workflow for manuscript and presentation assets.
- Complement Scientific Color Lab and other research-visualization tools.

---

## 03. Key features

| Module | What it does | 中文说明 |
|---|---|---|
| Page Range Export | Exports current slide, selected pages, ranges, or all slides | 支持当前页、指定页码、页码区间和全部页面导出 |
| High-DPI Control | Allows custom DPI settings, with 300 DPI as a common academic default | 支持自定义 DPI，常用 300 DPI 用于学术图件 |
| Multi-format Export | Supports PDF, PNG, JPG, BMP, and TIFF output | 支持 PDF、PNG、JPG、BMP、TIFF 导出 |
| PDF White-Border Crop | Exports native vector PDF without cropping, or lossless cropped PDF at the chosen DPI | PDF 不裁切时保留原生矢量内容；裁切时按 DPI 无损输出 |
| Save Path Selection | Lets users choose the output directory | 支持选择导出保存路径 |
| Academic Figure Workflow | Supports manuscript figures, presentations, posters, and graphical assets | 服务论文图件、学术汇报、海报和图形摘要 |
| PowerPoint Integration | Uses Microsoft Office Interop / add-in workflow | 基于 Microsoft Office Interop / 加载项流程 |

---

## 04. Product philosophy

ExportSlidesWithDPI_PPT follows four design principles:

1. **Resolution control** — figure export should not depend on PowerPoint defaults.
2. **Page-level precision** — users should export exactly the slides they need.
3. **Academic format awareness** — PNG, JPG, and TIFF outputs should fit manuscript and presentation workflows.
4. **Minimal interaction** — export should be faster than manual screenshot or repeated save-as operations.

---

## 05. Architecture

```text
ExportSlidesWithDPI_PPT
├── PowerPoint Add-in Layer
│   ├── UI controls
│   ├── page-range input
│   ├── DPI setting
│   └── format selection
├── Export Logic
│   ├── current slide export
│   ├── custom page range parser
│   ├── all-slide export
│   └── output path manager
└── Office Interop Layer
    ├── PowerPoint presentation object
    ├── slide object access
    └── image export API
```

---

## 06. Quick start

```bash
git clone https://github.com/groele/ExportSlidesWithDPI_PPT.git
cd ExportSlidesWithDPI_PPT
```

Development environment:

| Requirement | Recommendation |
|---|---|
| OS | Windows |
| PowerPoint | Microsoft PowerPoint desktop version |
| IDE | Visual Studio |
| Language | C# |
| API | Microsoft.Office.Interop.PowerPoint |

Build the project in Visual Studio and install/load the add-in according to the generated Office add-in package.

---

## 07. Recommended workflow

```text
Prepare slide figure → Choose page range
                     → Set DPI and image format
                     → Select output folder
                     → Export high-resolution images
                     → Use in manuscript / poster / presentation
```

Page-range examples:

| Input | Meaning |
|---|---|
| `0` | Export the currently selected slide |
| `1,3-5` | Export slide 1 and slides 3 to 5 |
| `all` | Export all slides |
| `2,4,6-8` | Export custom mixed ranges |

---

## 08. Roadmap

- [ ] Add detailed installation guide with screenshots
- [ ] Add batch export preset profiles
- [ ] Add automatic file naming templates
- [ ] Add transparent-background export notes where supported
- [ ] Add manuscript/journal DPI recommendations
- [ ] Add export log for reproducibility
- [ ] Add integration notes with Scientific Color Lab

---

## 09. Privacy and data ownership

ExportSlidesWithDPI_PPT runs locally with PowerPoint. Slide content and exported images remain on the user's machine unless manually shared or uploaded.

---

## 10. Related projects

- **PPT Presentation Timer** — academic presentation timing utility
- **Scientific Color Lab** — scientific color and visualization workspace
- **ManuGuide** — Microsoft Word manuscript formatting and style checker
- **ResearchFlow Companion** — research workflow operating system

---

## 11. License

MIT License.

Developed by **Shikun Hou / groele**.
