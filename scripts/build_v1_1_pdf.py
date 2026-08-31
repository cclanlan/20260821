#!/usr/bin/env python3
"""Merge chapter markdown files 00-12 into one printable HTML, then Chrome PDF."""

from __future__ import annotations

import subprocess
from pathlib import Path

import markdown

ROOT = Path("/workspace")
CH_DIR = ROOT / "金砖娱乐城商业计划书_V1.1_分章"
OUT_HTML = ROOT / "business-plan" / "金砖娱乐城商业计划书_V1.1_合并稿.html"
OUT_PDF = ROOT / "金砖娱乐城商业计划书_V1.1_合并稿.pdf"

CHAPTERS = [
    "00_封面与目录.md",
    "01_执行摘要.md",
    "02_项目背景与战略意义.md",
    "03_项目总览与核心定位.md",
    "04_建筑设计与功能规划.md",
    "05_市场分析与客源定位.md",
    "06_商业模式与收入模型.md",
    "07_投资规划与财务预测.md",
    "08_政府财政贡献.md",
    "09_法律政策与合规框架.md",
    "10_就业与社会责任.md",
    "11_实施时间表.md",
    "12_附录_修订摘要.md",
]

CSS = """
@page {
  size: A4;
  margin: 16mm 14mm 18mm 14mm;
}
html, body {
  font-family: "WenQuanYi Micro Hei", "Noto Sans CJK SC", "Microsoft YaHei", "SimSun", sans-serif;
  font-size: 11pt;
  line-height: 1.55;
  color: #1a1a1a;
  background: #fff;
}
h1 {
  font-size: 18pt;
  color: #0b1f3a;
  border-bottom: 2px solid #c5a572;
  padding-bottom: 6px;
  page-break-before: always;
  margin-top: 0;
}
h1:first-of-type { page-break-before: auto; }
h2 { font-size: 14pt; color: #0b1f3a; margin-top: 22px; }
h3 { font-size: 12pt; color: #6b5424; margin-top: 16px; }
h4 { font-size: 11pt; color: #6b5424; margin-top: 14px; }
p { margin: 0 0 8px 0; orphans: 3; widows: 3; }
strong { color: #0b1f3a; }
em { color: #555; }
hr { border: none; border-top: 1px solid #c5a572; margin: 18px 0; }
table {
  border-collapse: collapse;
  width: 100%;
  margin: 10px 0 14px 0;
  font-size: 9.5pt;
  page-break-inside: avoid;
}
th, td {
  border: 1px solid #c9c4b8;
  padding: 5px 7px;
  vertical-align: top;
  text-align: left;
}
th { background: #0b1f3a; color: #fff; }
tr:nth-child(even) td { background: #f7f4ee; }
.banner {
  background: #8b1e1e;
  color: #fff;
  text-align: center;
  padding: 6px 10px;
  font-size: 9.5pt;
  letter-spacing: 0.04em;
  margin: 0 0 16px 0;
}
.cover-meta { color: #444; }
"""

HTML_TMPL = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>金砖娱乐城综合开发项目商业计划书 V1.1 合并稿</title>
<style>{css}</style>
</head>
<body>
<div class="banner">机密文件 · 仅供政府审阅 &nbsp;|&nbsp; Strictly Confidential · For Government Review Only</div>
{body}
</body>
</html>
"""


def main() -> None:
    parts = []
    for name in CHAPTERS:
        path = CH_DIR / name
        if not path.exists():
            raise SystemExit(f"missing {path}")
        parts.append(path.read_text(encoding="utf-8").strip())
    md_text = "\n\n---\n\n".join(parts)
    body = markdown.markdown(
        md_text,
        extensions=["tables", "sane_lists", "nl2br"],
        output_format="html5",
    )
    html = HTML_TMPL.format(css=CSS, body=body)
    OUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    OUT_HTML.write_text(html, encoding="utf-8")

    cmd = [
        "google-chrome",
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--hide-scrollbars",
        "--no-pdf-header-footer",
        f"--print-to-pdf={OUT_PDF}",
        f"file://{OUT_HTML.resolve()}",
    ]
    subprocess.run(cmd, check=True, capture_output=True, text=True)
    if not OUT_PDF.exists() or OUT_PDF.stat().st_size < 10_000:
        raise SystemExit(f"PDF not generated: {OUT_PDF}")
    print("html", OUT_HTML, OUT_HTML.stat().st_size)
    print("pdf", OUT_PDF, OUT_PDF.stat().st_size)


if __name__ == "__main__":
    main()
