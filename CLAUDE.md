# CLAUDE.md

## Project Overview

VT PPTX Generator — converts structured Markdown files into Virginia Tech branded PowerPoint presentations using `python-pptx`.

## Quick Reference

- **Generate slides**: `./generate.sh weekly/<file>.md`
- **Create new report**: `./new.sh literature` or `./new.sh progress`
- **Install deps**: `pip3 install -r requirements.txt`
- **Template**: `vt_template.pptx` must exist in project root (not tracked by git — download from VT Brand Standards)

## Architecture

- `generate_pptx.py` — main script; parses Markdown → builds PPTX via VT template layouts
  - Layout 2 (Master 0): white background, corner logo — used for title + content slides
  - Layout 8 (Master 0): maroon background — used for section dividers
  - Placeholder indices: 0 = title, 10 = body text, 11 = date/content
- `generate.sh` — shell wrapper, auto-installs Python deps if missing
- `new.sh` — creates dated `.md` from template, replaces `YYYY-MM-DD` placeholder with today's date
- `templates/` — Markdown templates with YAML frontmatter + heading-based slide structure
- `weekly/` — user's weekly content files (gitignored)
- `output/` — generated `.pptx` files (gitignored)

## Markdown Format

- YAML frontmatter (`title`, `author`, `department`, `date`) → title slide
- `# Heading` → section divider slide (maroon)
- `## Heading` → content slide (white)
- `- bullet` → level 0 bullet; `  - bullet` → level 1 sub-bullet
- `**text**` → bold
- "Thank You" slide is auto-appended

## Important Constraints

- `vt_template.pptx` is copyrighted by Virginia Tech — never commit it to git
- Original template slides are hidden (`show="0"`), not deleted — deleting >12 slides causes PowerPoint repair prompts
- All presentation content must be in English
