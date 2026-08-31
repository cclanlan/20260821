#!/usr/bin/env python3
"""Build compact ZH + EN PDFs for BRICS Entertainment City."""

from __future__ import annotations

import subprocess
from pathlib import Path

import markdown

ROOT = Path("/workspace")
CH_DIR = ROOT / "金砖娱乐城商业计划书_V1.1_分章"
EN_MD = ROOT / "presentations" / "brics_entertainment_city_en.md"

ZH_CHAPTERS = [
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
    "12_人力资源成本分析.md",
    "13_附录_修订摘要.md",
]

CSS = """
@page { size: A4; margin: 10mm 11mm 11mm 11mm; }
html, body {
  font-family: "WenQuanYi Micro Hei", "Noto Sans CJK SC", "Microsoft YaHei", Calibri, sans-serif;
  font-size: 9.4pt;
  line-height: 1.32;
  color: #1a1a1a;
  background: #fff;
}
h1 {
  font-size: 13pt;
  color: #0b1f3a;
  border-bottom: 1.5px solid #c5a572;
  padding: 0 0 3px 0;
  margin: 10px 0 6px 0;
  page-break-before: auto;
  page-break-after: avoid;
}
h2 { font-size: 11pt; color: #0b1f3a; margin: 8px 0 3px 0; page-break-after: avoid; }
h3 { font-size: 10pt; color: #6b5424; margin: 6px 0 2px 0; page-break-after: avoid; }
h4 { font-size: 9.5pt; color: #6b5424; margin: 5px 0 2px 0; page-break-after: avoid; }
p { margin: 0 0 3.5px 0; orphans: 2; widows: 2; }
ul, ol { margin: 2px 0 4px 1.1em; padding: 0; }
li { margin: 0 0 1px 0; }
strong { color: #0b1f3a; }
em { color: #444; }
hr { border: none; border-top: 1px solid #ddd2b8; margin: 6px 0; }
table {
  border-collapse: collapse;
  width: 100%;
  margin: 3px 0 6px 0;
  font-size: 8pt;
  line-height: 1.25;
}
th, td {
  border: 1px solid #c9c4b8;
  padding: 2px 4px;
  vertical-align: top;
  text-align: left;
  overflow-wrap: anywhere;
  word-break: break-word;
}
th { background: #0b1f3a; color: #fff; font-weight: 600; }
tr:nth-child(even) td { background: #f7f4ee; }
.banner {
  background: #8b1e1e;
  color: #fff;
  text-align: center;
  padding: 3px 8px;
  font-size: 8pt;
  margin: 0 0 6px 0;
}
"""

HTML = """<!DOCTYPE html>
<html lang="{lang}">
<head>
<meta charset="utf-8">
<title>{title}</title>
<style>{css}</style>
</head>
<body>
<div class="banner">{banner}</div>
{body}
</body>
</html>
"""

REPLACES = [
    ("Bullion Entertainment City", "BRICS Entertainment City"),
    ("Jin Zhuan Entertainment City", "BRICS Entertainment City"),
    ("Bullion Grand Hotel", "BRICS Grand Hotel"),
    ("Bullion Tower", "BRICS Tower"),
    ("Bullion Entertainment (Sri Lanka)", "BRICS Entertainment (Sri Lanka)"),
    ("金砖娱乐城（Bullion", "金砖娱乐城（BRICS"),
    ("金砖娱乐城（Jin Zhuan", "金砖娱乐城（BRICS"),
]


def md_to_html(md_text: str, title: str, banner: str, lang: str) -> str:
    for a, b in REPLACES:
        md_text = md_text.replace(a, b)
    body = markdown.markdown(
        md_text,
        extensions=["tables", "sane_lists", "nl2br"],
        output_format="html5",
    )
    return HTML.format(lang=lang, title=title, css=CSS, banner=banner, body=body)


def chrome_pdf(html_path: Path, pdf_path: Path) -> None:
    if pdf_path.exists():
        pdf_path.unlink()
    cmd = [
        "google-chrome",
        "--headless",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--hide-scrollbars",
        "--no-pdf-header-footer",
        f"--print-to-pdf={pdf_path}",
        f"file://{html_path.resolve()}",
    ]
    try:
        subprocess.run(cmd, check=False, capture_output=True, text=True, timeout=90)
    except subprocess.TimeoutExpired:
        pass
    if not pdf_path.exists() or pdf_path.stat().st_size < 20_000:
        raise SystemExit(f"PDF failed: {pdf_path}")


def main() -> None:
    zh_md = "\n\n---\n\n".join(
        (CH_DIR / n).read_text(encoding="utf-8").strip() for n in ZH_CHAPTERS
    )
    en_md = EN_MD.read_text(encoding="utf-8")

    zh_html = ROOT / "business-plan" / "BRICS_Entertainment_City_ZH.html"
    en_html = ROOT / "business-plan" / "BRICS_Entertainment_City_EN.html"
    zh_pdf = ROOT / "BRICS_Entertainment_City_商业计划书_V1.2_中文.pdf"
    en_pdf = ROOT / "BRICS_Entertainment_City_Business_Plan_V1.2_EN.pdf"

    zh_html.write_text(
        md_to_html(
            zh_md,
            "BRICS Entertainment City 金砖娱乐城商业计划书 V1.2",
            "机密 · 仅供政府审阅  |  BRICS Entertainment City  |  Strictly Confidential",
            "zh-CN",
        ),
        encoding="utf-8",
    )
    en_html.write_text(
        md_to_html(
            en_md,
            "BRICS Entertainment City Business Plan V1.2",
            "Strictly Confidential · For Government Review Only  |  BRICS Entertainment City  |  机密",
            "en",
        ),
        encoding="utf-8",
    )

    chrome_pdf(zh_html, zh_pdf)
    chrome_pdf(en_html, en_pdf)

    # also refresh the old 合并稿 name so previous links still work
    combo = ROOT / "金砖娱乐城商业计划书_V1.1_合并稿.pdf"
    combo.write_bytes(zh_pdf.read_bytes())

    print("zh", zh_pdf, zh_pdf.stat().st_size)
    print("en", en_pdf, en_pdf.stat().st_size)


if __name__ == "__main__":
    main()
