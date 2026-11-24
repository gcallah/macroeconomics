#!/usr/bin/env python3
"""Build the macroeconomics static site from course material kept elsewhere.

Collects:
  * lecture decks (.pptx)  -> lectures/<slug>.html + lectures/<slug>/ images
  * quizzes and finals     -> quizzes/<slug>.html (via pandoc)
  * Excel models (.xlsx)   -> models/
and writes index.html linking them all.

Usage: python3 utils/build.py [--src DIR] [--models DIR] [--out DIR]
Needs: python-pptx (pip install python-pptx) and pandoc.
"""

import argparse
import hashlib
import html
import re
import shutil
import subprocess
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

HOME = Path.home()
DEFAULT_SRC = HOME / "Documents/CoursesTaught/Macroeconomics"
DEFAULT_MODELS = HOME / "GitProjects/MacroModels"

# Lecture decks, in course order: (source path relative to SRC, output slug).
# Left out on purpose: Chapter07/MAC_1e_Chap_7_M.pptx and MacroInContext/,
# which are the textbook publisher's slides.
DECKS = [
    ("Chapter01/Chapter1.pptx", "chapter01", "Chapter 1"),
    ("Chapter02/Chapter2.pptx", "chapter02", "Chapter 2"),
    ("Chapter03/Chapter3.pptx", "chapter03", "Chapter 3"),
    ("Chapter04/Chapter4.pptx", "chapter04", "Chapter 4"),
    ("Chapter05/Chapter5.pptx", "chapter05", "Chapter 5"),
    ("Chapter07/Chapter7.pptx", "chapter07", "Chapter 7"),
    ("Chapter07/Money.pptx", "chapter07-money", "Chapter 7"),
    ("Chapter09/Chapter09.pptx", "chapter09", "Chapter 9"),
    ("Chapter10/Chapter10.pptx", "chapter10", "Chapter 10"),
    ("Chapter11/Chapter11.pptx", "chapter11", "Chapter 11"),
    ("SchoolsOfThought/SchoolsOfThought.pptx", "schools-of-thought", "Supplement"),
]

# Slide titles that are only a chapter label; the deck's real name follows.
CHAPTER_TITLE_PREFIX = re.compile(r"^Chapter\s+\w+\s*:\s*", re.I)

# Quizzes and finals: .docx files in SRC whose names start with these.
# Excludes QuizTemplate, syllabi and other non-assessment documents.
QUIZ_PREFIXES = ("Quiz", "Final", "Macroeconomics Quiz", "MacroeconomicsQuiz")
QUIZ_EXCLUDE = {"QuizTemplate.docx"}

# Excel models: (file in MODELS dir, title, description).
MODELS = [
    ("RealVsNominal.xlsx", "Real versus nominal growth",
     "Generates problems for telling apart real growth, nominal growth and "
     "inflation in a two-good economy."),
    ("KeynesianCross.xlsx", "Keynesian cross",
     "Experiment with the Keynesian cross by setting autonomous consumption, "
     "intended investment and the marginal propensity to consume."),
    ("PPF.xlsx", "Production possibilities frontier",
     "Shows how the production possibilities frontier can change and how "
     "points along it are chosen."),
    ("LoanableFunds.xlsx", "Loanable funds", "The classical loanable funds market."),
    ("SupplyDemand.xlsx", "Supply and demand", "A basic supply and demand model."),
]

WEB_IMAGE_EXTS = {"png", "jpg", "jpeg", "gif", "svg", "webp"}
CONVERT_WITH_SIPS = {"tif", "tiff", "bmp"}

warnings = []


def warn(msg):
    warnings.append(msg)


def esc(text):
    return html.escape(text, quote=True)


SITE_TITLE = "Macroeconomics"
FONTS_URL = ("https://fonts.googleapis.com/css2?family=Archivo:wght@500;700;800"
             "&family=Source+Serif+4:ital,opsz,wght@0,8..60,400;0,8..60,600;"
             "1,8..60,400&display=swap")
CREDIT = ('Background: Adolph Menzel, <i>The Iron Rolling Mill (Modern Cyclopes)</i>, '
          '1872&ndash;75, Alte Nationalgalerie, Berlin. '
          '<a href="https://commons.wikimedia.org/wiki/File:Adolph_Menzel_-_Eisenwalzwerk_'
          '-_Google_Art_Project.jpg">Public domain, via Wikimedia Commons</a>.')

# Set in main() once style.css is known, so browsers refetch it after changes.
CSS_VERSION = ""


def page(title, body, root="", body_class="", hero=None):
    """Wrap body in the site chrome.

    root:  relative path from this page to the site root ("" or "../").
    hero:  optional (kicker, heading, subtitle) shown over the painting.
    """
    hero_html = ""
    if hero:
        kicker, heading, sub = hero
        hero_html = f"""<header class="hero">
  <div class="hero-inner">
    {f'<p class="kicker">{kicker}</p>' if kicker else ''}
    <h1>{heading}</h1>
    {f'<p class="sub">{sub}</p>' if sub else ''}
  </div>
</header>"""
    doc_title = SITE_TITLE if title == SITE_TITLE else f"{title} · {SITE_TITLE}"
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{esc(doc_title)}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link rel="stylesheet" href="{FONTS_URL}">
    <link rel="stylesheet" href="{root}style.css?v={CSS_VERSION}">
</head>
<body class="{body_class}">
<nav class="site-nav">
  <a class="brand" href="{root}index.html">{SITE_TITLE}</a>
  <span class="links">
    <a href="{root}index.html#lectures">Lectures</a>
    <a href="{root}index.html#models">Models</a>
    <a href="{root}index.html#quizzes">Quizzes</a>
  </span>
</nav>
{hero_html}
<main>
{body}
</main>
<footer class="site-footer">{CREDIT}</footer>
</body>
</html>
"""


# ---------------------------------------------------------------- slides

def run_html(run):
    text = esc(run.text).replace("\x0b", "<br>")
    if not text:
        return ""
    font = run.font
    if font.bold:
        text = f"<strong>{text}</strong>"
    if font.italic:
        text = f"<em>{text}</em>"
    addr = run.hyperlink.address if run.hyperlink else None
    if addr:
        text = f'<a href="{esc(addr)}">{text}</a>'
    return text


def text_frame_html(tf):
    """Render a text frame as nested lists keyed on paragraph level."""
    paras = [(p.level, "".join(run_html(r) for r in p.runs)) for p in tf.paragraphs]
    paras = [(lvl, t) for lvl, t in paras if t.strip()]
    if not paras:
        return ""
    out, depth = [], -1
    for lvl, text in paras:
        while depth < lvl:
            out.append("<ul>")
            depth += 1
        while depth > lvl:
            out.append("</li></ul>")
            depth -= 1
        if out and out[-1] not in ("<ul>",):
            out.append("</li>")
        out.append(f"<li>{text}")
    out.append("</li>")
    while depth >= 0:
        out.append("</ul>")
        depth -= 1
    return "\n".join(out)


def table_html(tbl):
    rows = []
    for r_i, row in enumerate(tbl.rows):
        tag = "th" if r_i == 0 else "td"
        cells = "".join(
            f"<{tag}>{esc(cell.text).replace(chr(10), '<br>')}</{tag}>"
            for cell in row.cells)
        rows.append(f"<tr>{cells}</tr>")
    return "<table>\n" + "\n".join(rows) + "\n</table>"


SERIES_COLORS = ["#1d4e89", "#c0392b", "#27864a", "#8e44ad"]


def cat_label(cat):
    s = str(cat).strip()
    return s[:-2] if s.endswith(".0") else s


def chart_html(chart):
    """Draw a line chart as inline SVG (the decks only use line charts)."""
    try:
        plot = chart.plots[0]
        cats = [cat_label(c) for c in plot.categories]
        series = [(s.name, list(s.values)) for s in plot.series]
    except Exception as e:  # noqa: BLE001
        warn(f"chart could not be read: {e}")
        return ""
    vals = [v for _, vs in series for v in vs if v is not None]
    if not vals or not cats:
        return ""
    lo, hi = min(vals + [0]), max(vals)
    if hi == lo:
        hi = lo + 1
    W, H, L, R, T, B = 640, 300, 56, 16, 16, 36
    n = len(cats)

    def x(i):
        return L + (W - L - R) * (i / max(n - 1, 1))

    def y(v):
        return T + (H - T - B) * (1 - (v - lo) / (hi - lo))

    parts = [f'<svg class="chart" viewBox="0 0 {W} {H}" role="img" '
             f'xmlns="http://www.w3.org/2000/svg">']
    for v in (lo, (lo + hi) / 2, hi):
        parts.append(f'<line x1="{L}" x2="{W - R}" y1="{y(v):.1f}" y2="{y(v):.1f}" class="grid"/>'
                     f'<text x="{L - 6}" y="{y(v) + 4:.1f}" text-anchor="end">{v:,.4g}</text>')
    if lo < 0 < hi:
        parts.append(f'<line x1="{L}" x2="{W - R}" y1="{y(0):.1f}" y2="{y(0):.1f}" class="zero"/>')
    # Label the first and last category plus where the category changes, sparsely.
    ticks, last = [], None
    for i, c in enumerate(cats):
        if c and c != last:
            ticks.append(i)
            last = c
    step = max(1, len(ticks) // 8)
    for i in ticks[::step]:
        parts.append(f'<text x="{x(i):.1f}" y="{H - B + 18}" text-anchor="middle">{esc(cats[i])}</text>')
    for k, (name, vs) in enumerate(series):
        color = SERIES_COLORS[k % len(SERIES_COLORS)]
        pts, segs = [], []
        for i, v in enumerate(vs):
            if v is None:
                if pts:
                    segs.append(pts)
                pts = []
            else:
                pts.append(f"{x(i):.1f},{y(v):.1f}")
        if pts:
            segs.append(pts)
        for seg in segs:
            parts.append(f'<polyline fill="none" stroke="{color}" stroke-width="2" '
                         f'points="{" ".join(seg)}"/>')
    parts.append("</svg>")
    legend = ""
    good_names = [(k, nm) for k, (nm, _) in enumerate(series)
                  if nm and nm not in ("Year",) and not nm.startswith("#")]
    if len(series) > 1 and good_names:
        legend = '<p class="legend">' + " ".join(
            f'<span style="color:{SERIES_COLORS[k % len(SERIES_COLORS)]}">&#9644; {esc(nm)}</span>'
            for k, nm in good_names) + "</p>"
    return "\n".join(parts) + legend


class Deck:
    def __init__(self, src, slug, outdir):
        self.src, self.slug = src, slug
        self.imgdir = outdir / slug
        self.n_images = 0

    def image_html(self, shape):
        img = shape.image
        ext = img.ext.lower()
        self.n_images += 1
        name = f"img{self.n_images:03d}"
        self.imgdir.mkdir(parents=True, exist_ok=True)
        if ext in CONVERT_WITH_SIPS:
            raw = self.imgdir / f"{name}.{ext}"
            raw.write_bytes(img.blob)
            dest = self.imgdir / f"{name}.png"
            subprocess.run(["sips", "-s", "format", "png", str(raw), "--out", str(dest)],
                           check=True, capture_output=True)
            raw.unlink()
        elif ext in WEB_IMAGE_EXTS:
            dest = self.imgdir / f"{name}.{ext}"
            dest.write_bytes(img.blob)
        else:
            warn(f"{self.src.name}: skipped .{ext} image (not viewable in browsers)")
            return ""
        alt = esc(getattr(shape, "name", "") or "")
        return f'<img src="{self.slug}/{dest.name}" alt="{alt}" loading="lazy">'

    def shape_html(self, shape, title_id):
        if shape.shape_id == title_id:
            return ""
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            return "\n".join(filter(None, (self.shape_html(s, title_id) for s in shape.shapes)))
        parts = []
        if hasattr(shape, "image"):
            try:
                parts.append(self.image_html(shape))
            except ValueError:
                pass  # placeholder with no picture in it
        if getattr(shape, "has_table", False) and shape.has_table:
            parts.append(table_html(shape.table))
        if getattr(shape, "has_chart", False) and shape.has_chart:
            parts.append(chart_html(shape.chart))
        if shape.has_text_frame:
            parts.append(text_frame_html(shape.text_frame))
        return "\n".join(filter(None, parts))


    def build(self):
        """Return (deck title, [(slide num, slide title)], slides html)."""
        prs = Presentation(self.src)
        if self.imgdir.exists():
            shutil.rmtree(self.imgdir)
        deck_title, toc, sections = None, [], []
        for num, slide in enumerate(prs.slides, 1):
            title_shape = slide.shapes.title
            if title_shape is not None and not title_shape.has_text_frame:
                title_shape = None
            title = title_shape.text_frame.text.strip() if title_shape is not None else ""
            title_id = title_shape.shape_id if title_shape is not None and title else None
            body = "\n".join(filter(None, (self.shape_html(s, title_id) for s in slide.shapes)))
            if not title:
                title = f"Slide {num}"
            title = re.sub(r"\s+", " ", title)
            if deck_title is None:
                deck_title = CHAPTER_TITLE_PREFIX.sub("", title)
            toc.append((num, title))
            notes = ""
            if slide.has_notes_slide:
                n = slide.notes_slide.notes_text_frame
                if n is not None and n.text.strip():
                    notes = f'<aside class="notes">{text_frame_html(n)}</aside>'
            sections.append(
                f'<section class="slide" id="s{num}">\n'
                f'<h2><a class="num" href="#s{num}">{num}</a> {esc(title)}</h2>\n'
                f'{body}\n{notes}\n</section>')
        return deck_title, toc, "\n".join(sections)


def pager(prev, nxt):
    """Previous/next lecture links. prev and nxt are (slug, title) or None."""
    def link(entry, cls, arrow):
        if not entry:
            return '<span></span>'
        slug, title = entry
        return (f'<a class="{cls}" href="{slug}.html"><span class="dir">{arrow}</span>'
                f'<span class="t">{esc(title)}</span></a>')
    return (f'<nav class="pager">{link(prev, "prev", "&larr; Previous")}'
            f'{link(nxt, "next", "Next &rarr;")}</nav>')


def build_lectures(src, out):
    outdir = out / "lectures"
    outdir.mkdir(parents=True, exist_ok=True)
    built = []
    for rel, slug, label in DECKS:
        path = src / rel
        if not path.exists():
            warn(f"missing deck: {path}")
            continue
        deck = Deck(path, slug, outdir)
        title, toc, body = deck.build()
        built.append((slug, label, title, toc, body))
        print(f"  {rel}: {len(toc)} slides, {deck.n_images} images")
    entries = []
    for i, (slug, label, title, toc, body) in enumerate(built):
        prev = built[i - 1][:3:2] if i > 0 else None
        nxt = built[i + 1][:3:2] if i + 1 < len(built) else None
        contents = "\n".join(f'<li><a href="#s{n}">{esc(t)}</a></li>' for n, t in toc)
        content = f"""{pager(prev, nxt)}
<details class="contents">
<summary>Contents &middot; {len(toc)} slides</summary>
<ol>
{contents}
</ol>
</details>
{body}
{pager(prev, nxt)}"""
        (outdir / f"{slug}.html").write_text(
            page(title, content, root="../", body_class="lecture",
                 hero=(esc(label), esc(title), f"{len(toc)} slides")),
            encoding="utf-8")
        entries.append((f"lectures/{slug}.html", label, title, len(toc)))
    return entries


# ---------------------------------------------------------------- quizzes

def slugify(name):
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


QUIZ_NAME = re.compile(
    r"^(?:Macroeconomics\s*)?(?P<kind>Quiz|Final)\s*(?P<num>\d*)(?P<ver>[a-z]?)"
    r"(?P<makeup>Makeup)?(?:[-_](?P<year>\d{2,4})(?P<sem>S)?)?$")


def describe_quiz(stem):
    """Map a file name to (group, label), e.g. Quiz1-16S -> ("Quiz 1", "Spring 2016")."""
    m = QUIZ_NAME.match(stem)
    if not m:
        return ("Other", re.sub(r"(?<=[a-z])(?=[A-Z])", " ", stem))
    group = "Final exam" if m["kind"] == "Final" else f"Quiz {m['num']}"
    bits = []
    if m["year"]:
        year = m["year"] if len(m["year"]) == 4 else "20" + m["year"]
        bits.append(("Spring " if m["sem"] else "") + year)
    if m["ver"]:
        bits.append(f"Version {m['ver'].upper()}")
    if m["makeup"]:
        bits.append("Make-up")
    return (group, ", ".join(bits) or "Undated")


def build_quizzes(src, out):
    outdir = out / "quizzes"
    if outdir.exists():
        shutil.rmtree(outdir)
    outdir.mkdir(parents=True)
    entries, seen = [], {}
    files = sorted(p for p in src.glob("*.docx")
                   if p.name.startswith(QUIZ_PREFIXES) and p.name not in QUIZ_EXCLUDE)
    for path in files:
        slug = slugify(path.stem)
        media = outdir / slug
        res = subprocess.run(
            ["pandoc", str(path), "-t", "html5", f"--extract-media={media}"],
            capture_output=True, text=True, check=True, cwd=outdir)
        body = res.stdout.replace(f'src="{media}/', f'src="{slug}/')
        group, label = describe_quiz(path.stem)
        seen[(group, label)] = seen.get((group, label), 0) + 1
        if seen[(group, label)] > 1:
            label = f"{label} ({seen[(group, label)]})"
        title = f"{group} ({label})" if group != "Other" else label
        (outdir / f"{slug}.html").write_text(
            page(title, f'<article class="paper">\n{body}\n</article>', root="../",
                 body_class="quiz", hero=("Quizzes and exams", esc(group if group != "Other" else label),
                                          esc(label) if group != "Other" else "")),
            encoding="utf-8")
        entries.append((f"quizzes/{slug}.html", group, label))
    print(f"  {len(entries)} quizzes/finals")
    return entries


# ---------------------------------------------------------------- models

def build_models(models, out):
    outdir = out / "models"
    outdir.mkdir(parents=True, exist_ok=True)
    entries = []
    for fname, title, desc in MODELS:
        path = models / fname
        if not path.exists():
            warn(f"missing model: {path}")
            continue
        shutil.copy2(path, outdir / fname)
        entries.append((f"models/{fname}", title, desc))
    print(f"  {len(entries)} models")
    return entries


# ---------------------------------------------------------------- index

def group_order(group):
    if group.startswith("Quiz "):
        return (0, int(group.split()[1] or 0))
    return (1, 0) if group == "Final exam" else (2, 0)


def build_index(out, lectures, quizzes, models):
    lec = "\n".join(
        f'<li><a class="card" href="{h}"><span class="label">{esc(lb)}</span>'
        f'<span class="title">{esc(t)}</span><span class="meta">{n} slides</span></a></li>'
        for h, lb, t, n in lectures)
    md = "\n".join(
        f'<li><a class="card" href="{h}" download><span class="label">Excel model</span>'
        f'<span class="title">{esc(t)}</span><span class="desc">{esc(d)}</span>'
        f'<span class="meta">Download .xlsx</span></a></li>'
        for h, t, d in models)
    groups = {}
    for h, g, lb in quizzes:
        groups.setdefault(g, []).append((lb, h))
    qz = "\n".join(
        f'<li class="quiz-group"><span class="title">{esc(g)}</span><span class="chips">'
        + "".join(f'<a href="{h}">{esc(lb)}</a>' for lb, h in sorted(groups[g]))
        + "</span></li>"
        for g in sorted(groups, key=group_order))
    body = f"""<section id="lectures">
<h2>Lectures</h2>
<ul class="cards">
{lec}
</ul>
</section>
<section id="models">
<h2>Models</h2>
<p class="lede">Spreadsheets for experimenting with core macroeconomic models.</p>
<ul class="cards">
{md}
</ul>
</section>
<section id="quizzes">
<h2>Quizzes and exams</h2>
<p class="lede">Past quizzes and finals, grouped by quiz.</p>
<ul class="quiz-list">
{qz}
</ul>
</section>"""
    (out / "index.html").write_text(
        page(SITE_TITLE, body, body_class="home",
             hero=("A course in", SITE_TITLE,
                   "Lectures, models and exams on output, employment, money and fluctuations.")),
        encoding="utf-8")


def main():
    global CSS_VERSION
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--src", type=Path, default=DEFAULT_SRC)
    ap.add_argument("--models", type=Path, default=DEFAULT_MODELS)
    ap.add_argument("--out", type=Path, default=Path(__file__).resolve().parent.parent)
    args = ap.parse_args()
    for p in (args.src, args.models):
        if not p.is_dir():
            sys.exit(f"source directory not found: {p}")
    CSS_VERSION = hashlib.sha1((args.out / "style.css").read_bytes()).hexdigest()[:8]
    print("Lectures:")
    lectures = build_lectures(args.src, args.out)
    print("Quizzes:")
    quizzes = build_quizzes(args.src, args.out)
    print("Models:")
    models = build_models(args.models, args.out)
    build_index(args.out, lectures, quizzes, models)
    for w in warnings:
        print("warning:", w, file=sys.stderr)


if __name__ == "__main__":
    main()
