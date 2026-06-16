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
  --highlight: #f59e0b;
  --highlight-soft: #fff7ed;
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
  color: var(--ink);
  text-decoration: none;
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
  gap: 10px;
  flex-wrap: wrap;
}

nav a,
.button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 40px;
  padding: 8px 14px;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: var(--panel);
  color: var(--ink);
  text-decoration: none;
  font-size: 15px;
}

nav a:hover,
.button:hover {
  border-color: var(--accent);
}

nav a[aria-current="page"] {
  border-color: var(--accent);
  color: var(--accent-strong);
}

nav a.feature-link {
  border-color: var(--highlight);
  background: var(--highlight-soft);
  color: #7c2d12;
  font-weight: 700;
}

nav a.source-link {
  border-color: var(--accent);
  background: #ecfeff;
  color: var(--accent-strong);
  font-weight: 700;
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

.featured {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 18px;
  align-items: center;
  margin: 28px 0 10px;
  padding: 22px;
  border: 2px solid var(--highlight);
  border-radius: 8px;
  background: var(--highlight-soft);
  color: #431407;
  text-decoration: none;
}

.featured h2 {
  margin: 0;
  font-size: clamp(26px, 4vw, 38px);
}

.featured p {
  margin: 6px 0 0;
  color: #7c2d12;
}

.featured .button {
  border-color: var(--highlight);
  background: #ffffff;
  color: #7c2d12;
  font-weight: 700;
}

.grid {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 16px;
  margin-top: 18px;
}

.card {
  display: grid;
  align-content: start;
  gap: 12px;
  min-height: 190px;
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
  font-size: 22px;
}

.card p {
  color: var(--muted);
}

.source-card {
  border-color: var(--accent);
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

.figure-page,
.media-page {
  display: grid;
  gap: 22px;
}

.full-image,
.media-player {
  width: 100%;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: var(--panel);
}

.full-image {
  height: auto;
}

.media-player {
  max-height: 72vh;
}

footer {
  padding: 28px 0;
  border-top: 1px solid var(--line);
  color: var(--muted);
  background: var(--panel);
  font-size: 14px;
}

@media (max-width: 900px) {
  .grid {
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }
}

@media (max-width: 760px) {
  .mast {
    align-items: flex-start;
    flex-direction: column;
    padding: 16px 0;
  }

  .featured {
    grid-template-columns: 1fr;
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
        ("index.html", "1. 首頁", "首頁", ""),
        ("01-gcp-tsai-pdf.html", "2. 簡報", "簡報", ""),
        ("02-gcp-tsai-word.html", "3. 文章", "文章", ""),
        ("03-gcp-ban02m-image.html", "4. 圖片", "圖片", ""),
        ("04-speech-summary.html", "5. 語音", "語音", "feature-link"),
        ("https://youtu.be/rx4NNViz0e0", "6. 來源影片(Dr. Tsai)", "來源影片", "source-link"),
    ]
    nav = "\n".join(
        f'<a href="{href}" class="{css_class}"{" aria-current=\"page\"" if key == active else ""}{ " target=\"_blank\" rel=\"noopener\"" if href.startswith("https://") else ""}>{label}</a>'
        for href, label, key, css_class in links
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
        <span>線上閱讀資料</span>
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
        paragraphs.append(text)
    return paragraphs


def article_html(docx_path: Path) -> tuple[str, str]:
    paragraphs = clean_paragraphs(docx_path)
    title = paragraphs[1] if len(paragraphs) > 1 else "文章"
    parts: list[str] = [
        '<section class="hero">',
        '<div class="kicker">文章</div>',
        f"<h1>{html.escape(title)}</h1>",
        '<div class="actions"><a class="button" href="assets/02_GCP_Tsai.docx">下載原始檔</a></div>',
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
    for stale in (
        SITE / "01-gcp-tasi-pdf.html",
        SITE / "02-gcp-tasi-word.html",
        ASSETS / "01_GCP_TASI.pdf",
        ASSETS / "02_GCP_TASI.docx",
    ):
        if stale.exists():
            stale.unlink()

    pdf = SOURCE / "01_GCP_TASI.pdf"
    docx = SOURCE / "02_GCP_TASI.docx"
    image = SOURCE / "03_GCP_ban02m.png"
    speech = SOURCE / "04_AI_Speech_Summary.mp4"

    asset_names = {
        pdf: "01_GCP_Tsai.pdf",
        docx: "02_GCP_Tsai.docx",
        image: "03_GCP_ban02m.png",
        speech: "04_AI_Speech_Summary.mp4",
    }

    for src, name in asset_names.items():
        shutil.copy2(src, ASSETS / name)

    write(SITE / "styles.css", CSS + "\n")

    pdf_pages = len(PdfReader(str(pdf)).pages)
    with Image.open(image) as img:
        image_width, image_height = img.size

    speech_mb = speech.stat().st_size / 1024 / 1024

    index_body = f"""
<section class="hero">
  <div class="kicker">線上閱讀版</div>
  <h1>資料整理</h1>
  <p class="meta">已整理成可直接瀏覽的網頁，包含簡報、文章、圖片與語音摘要。</p>
</section>
<a class="featured" href="04-speech-summary.html">
  <div>
    <div class="kicker">新增內容</div>
    <h2>語音</h2>
    <p>語音摘要已加入，可直接播放，也可下載保存。</p>
  </div>
  <span class="button">前往語音</span>
</a>
<section class="grid" aria-label="資料列表">
  <a class="card" href="01-gcp-tsai-pdf.html">
    <span class="kicker">1</span>
    <h2>簡報</h2>
    <p>{pdf_pages} 頁簡報，可線上閱讀並下載原始檔。</p>
  </a>
  <a class="card" href="02-gcp-tsai-word.html">
    <span class="kicker">2</span>
    <h2>文章</h2>
    <p>已轉成適合手機與電腦閱讀的文章頁。</p>
  </a>
  <a class="card" href="03-gcp-ban02m-image.html">
    <span class="kicker">3</span>
    <h2>圖片</h2>
    <p>原圖尺寸 {image_width} x {image_height}，可瀏覽或下載。</p>
  </a>
  <a class="card" href="04-speech-summary.html">
    <span class="kicker">4</span>
    <h2>語音</h2>
    <p>語音摘要約 {speech_mb:.1f} MB，支援線上播放。</p>
  </a>
  <a class="card source-card" href="https://youtu.be/rx4NNViz0e0" target="_blank" rel="noopener">
    <span class="kicker">5</span>
    <h2>來源影片</h2>
    <p>蔡博士原始語音檔來源影片，將開啟 YouTube 連結。</p>
  </a>
</section>
"""
    write(SITE / "index.html", page("資料整理", index_body, "首頁"))

    pdf_body = f"""
<section class="hero">
  <div class="kicker">簡報</div>
  <h1>簡報線上閱讀</h1>
  <p class="meta">共 {pdf_pages} 頁。若瀏覽器沒有顯示內容，可使用下方按鈕開啟或下載原始檔。</p>
  <div class="actions">
    <a class="button" href="assets/01_GCP_Tsai.pdf">開啟檔案</a>
    <a class="button" href="assets/01_GCP_Tsai.pdf" download>下載檔案</a>
  </div>
</section>
<object class="reader" data="assets/01_GCP_Tsai.pdf" type="application/pdf">
  <p>瀏覽器無法內嵌簡報。請改用 <a href="assets/01_GCP_Tsai.pdf">檔案連結</a> 開啟。</p>
</object>
"""
    write(SITE / "01-gcp-tsai-pdf.html", page("簡報", pdf_body, "簡報"))

    article_title, article_body = article_html(docx)
    write(SITE / "02-gcp-tsai-word.html", page(article_title, article_body, "文章"))

    image_body = f"""
<section class="hero">
  <div class="kicker">圖片</div>
  <h1>圖片瀏覽</h1>
  <p class="meta">原圖尺寸 {image_width} x {image_height}。</p>
  <div class="actions">
    <a class="button" href="assets/03_GCP_ban02m.png">開啟原圖</a>
    <a class="button" href="assets/03_GCP_ban02m.png" download>下載圖片</a>
  </div>
</section>
<section class="figure-page">
  <img class="full-image" src="assets/03_GCP_ban02m.png" alt="圖片">
</section>
"""
    write(SITE / "03-gcp-ban02m-image.html", page("圖片", image_body, "圖片"))

    speech_body = f"""
<section class="hero">
  <div class="kicker">語音</div>
  <h1>語音摘要</h1>
  <p class="meta">可直接播放語音摘要；若瀏覽器無法播放，請下載原始檔。</p>
  <div class="actions">
    <a class="button" href="assets/04_AI_Speech_Summary.mp4">開啟檔案</a>
    <a class="button" href="assets/04_AI_Speech_Summary.mp4" download>下載語音</a>
  </div>
</section>
<section class="media-page">
  <video class="media-player" controls preload="metadata">
    <source src="assets/04_AI_Speech_Summary.mp4" type="video/mp4">
    您的瀏覽器無法播放此檔案，請使用下載按鈕取得原始檔。
  </video>
</section>
"""
    write(SITE / "04-speech-summary.html", page("語音", speech_body, "語音"))

    readme = """# 線上閱讀資料

這個資料夾已整理成 GitHub Pages 可發布的靜態網站。

## 內容

- `docs/index.html`: 首頁
- `docs/01-gcp-tsai-pdf.html`: 簡報線上閱讀頁
- `docs/02-gcp-tsai-word.html`: 文章頁
- `docs/03-gcp-ban02m-image.html`: 圖片瀏覽頁
- `docs/04-speech-summary.html`: 語音頁
- 來源影片(Dr. Tsai): https://youtu.be/rx4NNViz0e0
- `docs/assets/`: 原始檔案

## GitHub Pages 設定

GitHub Pages 使用 `main` branch 的 `/docs` folder 發布。
"""
    write(ROOT / "README.md", readme)


if __name__ == "__main__":
    main()
