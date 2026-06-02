# VT Slides

Generate branded Virginia Tech PowerPoint presentations from Markdown.

Write your content in a simple `.md` file, run one command, and get a `.pptx` with VT's official template — maroon dividers, corner logos, and all.

Built for weekly lab meetings, but works for any VT presentation.

## Features

- Markdown to PPTX — no manual slide formatting
- Uses VT's official widescreen PowerPoint template
- Two built-in report templates: **literature report** and **progress report**
- Section dividers (maroon background) and content slides (white with VT branding)
- Bold text support via `**bold**` syntax
- Auto-generated title slide and closing "Thank You" slide

## Prerequisites

- Python 3.6+
- macOS / Linux (shell scripts use `bash` and `sed`)

## Setup

### 1. Clone the repository

```bash
git clone https://github.com/liaochengjiaao-oliver/vt-slides.git
cd vt-slides
```

### 2. Install Python dependencies

```bash
pip3 install -r requirements.txt
```

### 3. Download the VT template

The official VT PowerPoint template is copyrighted by Virginia Tech and cannot be redistributed. Download it yourself:

1. Download from [Virginia Tech Brand Standards](https://brand.vt.edu/content/dam/brand_vt_edu/downloads/ppt/VT_PresentationTemplate_widescreen_PowerPoint.pptx.zip)
2. Extract the `.zip` file
3. Rename the `.pptx` file to `vt_template.pptx` and place it in the project root

```bash
# Or use the command line:
curl -L -o vt_template.zip "https://brand.vt.edu/content/dam/brand_vt_edu/downloads/ppt/VT_PresentationTemplate_widescreen_PowerPoint.pptx.zip"
unzip vt_template.zip
mv *.pptx vt_template.pptx
rm vt_template.zip
```

## Usage

### Create a new report

```bash
./new.sh literature   # Create a literature report
./new.sh progress     # Create a progress report
```

This creates a dated `.md` file in `weekly/` (e.g., `weekly/2026-06-02-literature.md`).

### Edit your content

Open the generated `.md` file and replace the placeholder content with your own.

### Generate slides

```bash
./generate.sh weekly/2026-06-02-literature.md
```

Your `.pptx` will be saved to `output/`.

## Markdown Format

```markdown
---
title: "Your Presentation Title"
author: "Your Name"
department: "Your Department"
date: "2026-06-02"
---

## Paper Information
- **Title:** Full Paper Title
- **Authors:** Author1, Author2
- **Venue:** Conference Name, 2026

# Background

## Research Background
- Core problem in this field
- Limitations of existing methods
  - Sub-point with more detail

## Motivation
- Key gap identified
- Core insight

# Method

## Method Overview
- Overall framework
- Key contributions
```

### Syntax

| Markdown | Result |
|----------|--------|
| `# Heading` | Section divider slide (maroon background) |
| `## Heading` | Content slide (white background, VT branding) |
| `- item` | Bullet point |
| `  - item` | Sub-bullet (indent 2 spaces) |
| `**text**` | Bold text |
| YAML frontmatter | Title slide (auto-generated) |

A "Thank You / Q&A" slide is automatically appended.

## Project Structure

```
.
├── generate_pptx.py      # Markdown to PPTX generation script
├── generate.sh            # Shell wrapper (auto-installs deps)
├── new.sh                 # Create new weekly report from template
├── vt_template.pptx       # VT template (download separately, not tracked)
├── requirements.txt       # Python dependencies
├── templates/
│   ├── literature_report.md   # Literature report template
│   └── progress_report.md    # Progress report template
├── weekly/                # Your weekly .md files (not tracked)
└── output/                # Generated .pptx files (not tracked)
```

## Customization

- **Layouts**: The script uses Layout 2 (content) and Layout 8 (divider) from the VT template. Adjust `LAYOUT_TITLE`, `LAYOUT_DIVIDER`, and `LAYOUT_CONTENT` in `generate_pptx.py` if your template version differs.
- **Colors**: VT Maroon (`#861F41`), Burnt Orange (`#E87722`), and Dark (`#3A3C3D`) are defined at the top of `generate_pptx.py`.
- **Templates**: Edit or add `.md` files in `templates/` to create new report types.

## License

[MIT](LICENSE)

## Acknowledgments

- PowerPoint template by [Virginia Tech Brand Standards](https://brand.vt.edu/)
- Built with [python-pptx](https://python-pptx.readthedocs.io/)
