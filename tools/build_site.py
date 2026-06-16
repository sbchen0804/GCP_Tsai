from __future__ import annotations

import html
import re
import shutil
from pathlib import Path

from docx import Document
from PIL import Image
from pypdf import PdfReader


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "INPUT_source"
SITE = ROOT / "docs"
ASSETS = SITE / "assets"


CSS = """
:root {
  color-scheme: light;
  --ink: #17212b;
  --muted: #5d6b78;
  --line: #d9e1e8;
  --paper: #fbfcfd;
  --panel: #ffffff;
  --accent: #0f766e;
  --accent-strong: #134e4a;
}

* {
  box-sizing: border-box;
}

body {
  margin: 0;
  background: var(--paper);
  color: var(--ink);
  font-family: "Noto Sans TC", "Microsoft JhengHei", "PingFang TC", Arial, sans-serif;
  line-height: 1.72;
}

a {
  color: var(--accent-strong);
}

.wrap {
  width: min(1080px, calc(100% - 32px));
  margin: 0 auto;
}

header.site {
  border-bottom: 1px solid var(--line);
  background: var(--panel);
}

.mast {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 24px;
  min-height: 76px;
}

.brand {
  display: grid;
  gap: 2px;
}

.brand strong {
  font-size: 20px;
  letter-spacing: 0;
}

.brand span,
.meta,
.kicker {
  color: var(--muted);
  font-size: 14px;
}

nav {
  display: flex;
  gap: 14px;
  flex-wrap: wrap;
}

nav a,
.button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 40px;
  padding: 8px 13px;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: var(--panel);
  color: var(--ink);
  text-decoration: none;
  font-size: 14px;
}

nav a:hover,
.button:hover {
  border-color: var(--accent);
}

main {
  padding: 42px 0 64px;
}

.hero {
  display: grid;
  gap: 14px;
  padding-bottom: 28px;
  border-bottom: 1px solid var(--line);
}

h1 {
  margin: 0;
  max-width: 880px;
  font-size: clamp(30px, 5vw, 52px);
  line-height: 1.14;
  letter-spacing: 0;
}

h2 {
  margin: 34px 0 12px;
  font-size: 24px;
  line-height: 1.3;
  letter-spacing: 0;
}

h3 {
  margin: 26px 0 8px;
  font-size: 20px;
  line-height: 1.35;
  letter-spacing: 0;
}

p {
  margin: 0 0 16px;
}

.grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 16px;
  margin-top: 24px;
}

.card {
  display: grid;
  align-content: start;
  gap: 12px;
  min-height: 220px;
  padding: 20px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: var(--panel);
  text-decoration: none;
  color: var(--ink);
}

.card:hover {
  border-color: var(--accent);
}

.card h2 {
  margin: 0;
  font-size: 21px;
}

.card p {
  color: var(--muted);
}

.actions {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  margin-top: 20px;
}

.reader {
  width: 100%;
  height: min(78vh, 860px);
  min-height: 540px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: var(--panel);
}

.article {
  max-width: 820px;
  padding-top: 22px;
}

.article p {
  font-size: 18px;
}

.article .lead {
  font-size: 20px;
  color: #263746;
}

.figure-page {
  display: grid;
  gap: 22px;
}

.full-image {
  width: 100%;
  height: auto;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: var(--panel);
}

footer {
  padding: 28px 0;
  border-top: 1px solid var(--line);
  color: var(--muted);
  background: var(--panel);
  font-size: 14px;
}

@media (max-width: 760px) {
  .mast {
    align-items: flex-start;
    flex-direction: column;
    padding: 16px 0;
  }

  .grid {
    grid-template-columns: 1fr;
  }

  .reader {
    min-height: 420px;
    height: 70vh;
  }
}
""".strip()


def page(title: str, body: str, active: str = "") -> str:
    links = [
        ("index.html", "首頁"),
        ("01-gcp-tasi-pdf.html", "PDF 簡報"),
        ("02-gcp-tasi-word.html", "Word 文章"),
        ("03-gcp-ban02m-image.html", "圖片"),
    ]
    nav = "\n".join(
        f'<a href="{href}"{" aria-current=\"page\"" if label == active else ""}>{label}</a>'
        for href, label in links
    )
    return f"""<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <link rel="stylesheet" href="styles.css">
</head>
<body>
  <header class="site">
    <div class="wrap mast">
      <a class="brand" href="index.html" aria-label="回首頁">
        <strong>蔡憶雲 中醫博士</strong>
        <span>GCP TASI 線上閱讀資料</span>
      </a>
      <nav aria-label="主要導覽">
        {nav}
      </nav>
    </div>
  </header>
  <main>
    <div class="wrap">
      {body}
    </div>
  </main>
  <footer>
    <div class="wrap">本頁面為 GitHub Pages 靜態網站版本，保留原始檔供下載與備份。</div>
  </footer>
</body>
</html>
"""


def clean_paragraphs(docx_path: Path) -> list[str]:
    doc = Document(docx_path)
    paragraphs: list[str] = []
    carry = ""
    for paragraph in doc.paragraphs:
        text = " ".join(paragraph.text.split())
        if not text:
            continue
        if text in {"。", "！", "，", "、"} and paragraphs:
            paragraphs[-1] += text
            continue
        if text[0] in "。！，、；：" and paragraphs:
            paragraphs[-1] += text
            continue
        if carry:
            text = carry + text
            carry = ""
        paragraphs.append(text)
    return paragraphs


def article_html(docx_path: Path) -> tuple[str, str]:
    paragraphs = clean_paragraphs(docx_path)
    title = paragraphs[1] if len(paragraphs) > 1 else "02_GCP_TASI"
    parts: list[str] = [
        '<section class="hero">',
        '<div class="kicker">02_GCP_TASI Word 轉 HTML</div>',
        f"<h1>{html.escape(title)}</h1>",
        '<div class="actions"><a class="button" href="assets/02_GCP_TASI.docx">下載原始 Word 檔</a></div>',
        "</section>",
        '<article class="article">',
    ]
    for index, text in enumerate(paragraphs):
        escaped = html.escape(text)
        if index == 0:
            parts.append(f'<p class="kicker">{escaped}</p>')
        elif index == 1:
            continue
        elif len(text) <= 24 and not text[0].isdigit() and text[-1] not in "。！？":
            parts.append(f"<h2>{escaped}</h2>")
        elif re.match(r"^\d+\.\s+", text):
            match = re.match(r"^(\d+\.\s+[^ ]+)\s+(.+)$", text)
            if match:
                parts.append(f"<h3>{html.escape(match.group(1))}</h3>")
                parts.append(f"<p>{html.escape(match.group(2))}</p>")
            else:
                parts.append(f"<h3>{escaped}</h3>")
        elif index == 2:
            parts.append(f'<p class="lead">{escaped}</p>')
        else:
            parts.append(f"<p>{escaped}</p>")
    parts.append("</article>")
    return title, "\n".join(parts)


def write(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8", newline="\n")


def main() -> None:
    SITE.mkdir(exist_ok=True)
    ASSETS.mkdir(parents=True, exist_ok=True)

    pdf = SOURCE / "01_GCP_TASI.pdf"
    docx = SOURCE / "02_GCP_TASI.docx"
    image = SOURCE / "03_GCP_ban02m.png"

    for src in (pdf, docx, image):
        shutil.copy2(src, ASSETS / src.name)

    write(SITE / "styles.css", CSS + "\n")

    pdf_pages = len(PdfReader(str(pdf)).pages)
    with Image.open(image) as img:
        image_width, image_height = img.size

    index_body = f"""
<section class="hero">
  <div class="kicker">線上閱讀版</div>
  <h1>GCP TASI 資料整理</h1>
  <p class="meta">已整理成 GitHub Pages 可直接瀏覽的 HTML 文件，包含 PDF 簡報、Word 文章與圖片檔。</p>
</section>
<section class="grid" aria-label="資料列表">
  <a class="card" href="01-gcp-tasi-pdf.html">
    <span class="kicker">01_GCP_TASI</span>
    <h2>PDF 簡報</h2>
    <p>{pdf_pages} 頁 PDF，使用線上閱讀器嵌入，並提供原始 PDF 下載。</p>
  </a>
  <a class="card" href="02-gcp-tasi-word.html">
    <span class="kicker">02_GCP_TASI</span>
    <h2>Word 轉 HTML</h2>
    <p>已將 Word 內容轉成適合手機與桌機閱讀的 HTML 文章頁。</p>
  </a>
  <a class="card" href="03-gcp-ban02m-image.html">
    <span class="kicker">03_GCP_ban02m</span>
    <h2>圖片檔</h2>
    <p>原圖尺寸 {image_width} x {image_height}，可直接瀏覽或下載。</p>
  </a>
</section>
"""
    write(SITE / "index.html", page("GCP TASI 資料整理", index_body, "首頁"))

    pdf_body = f"""
<section class="hero">
  <div class="kicker">01_GCP_TASI PDF</div>
  <h1>PDF 簡報線上閱讀</h1>
  <p class="meta">共 {pdf_pages} 頁。若瀏覽器沒有顯示 PDF，可使用下方按鈕開啟或下載原始檔。</p>
  <div class="actions">
    <a class="button" href="assets/01_GCP_TASI.pdf">開啟 PDF</a>
    <a class="button" href="assets/01_GCP_TASI.pdf" download>下載 PDF</a>
  </div>
</section>
<object class="reader" data="assets/01_GCP_TASI.pdf" type="application/pdf">
  <p>瀏覽器無法內嵌 PDF。請改用 <a href="assets/01_GCP_TASI.pdf">PDF 連結</a> 開啟。</p>
</object>
"""
    write(SITE / "01-gcp-tasi-pdf.html", page("01_GCP_TASI PDF", pdf_body, "PDF 簡報"))

    article_title, article_body = article_html(docx)
    write(SITE / "02-gcp-tasi-word.html", page(article_title, article_body, "Word 文章"))

    image_body = f"""
<section class="hero">
  <div class="kicker">03_GCP_ban02m PNG</div>
  <h1>圖片線上瀏覽</h1>
  <p class="meta">原圖尺寸 {image_width} x {image_height}。</p>
  <div class="actions">
    <a class="button" href="assets/03_GCP_ban02m.png">開啟原圖</a>
    <a class="button" href="assets/03_GCP_ban02m.png" download>下載 PNG</a>
  </div>
</section>
<section class="figure-page">
  <img class="full-image" src="assets/03_GCP_ban02m.png" alt="03_GCP_ban02m 圖檔">
</section>
"""
    write(SITE / "03-gcp-ban02m-image.html", page("03_GCP_ban02m 圖片", image_body, "圖片"))

    readme = """# GCP TASI 線上閱讀資料

這個資料夾已整理成 GitHub Pages 可發布的靜態網站。

## 內容

- `site/index.html`: 首頁
- `site/01-gcp-tasi-pdf.html`: PDF 簡報線上閱讀頁
- `site/02-gcp-tasi-word.html`: Word 轉換後的 HTML 文章頁
- `site/03-gcp-ban02m-image.html`: 圖片瀏覽頁
- `site/assets/`: 原始 PDF、Word、PNG 檔案

## GitHub Pages 設定

上傳到 GitHub 後，可在 repository 的 Settings -> Pages 選擇從 `main` branch 的 `/site` folder 發布。
"""
    write(ROOT / "README.md", readme)


if __name__ == "__main__":
    main()
