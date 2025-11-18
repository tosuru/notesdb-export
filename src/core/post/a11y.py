
"""
HTML accessibility enhancement utilities.
Uses lxml.html to post-process rendered HTML and add common a11y affordances.
Provided functions are intentionally conservative to avoid layout regressions.
"""
from __future__ import annotations
from typing import Optional
from lxml import html as _html


def enhance_accessibility(html: str, lang: str = "ja") -> str:
    """
    Apply a bundle of small, safe a11y upgrades:
    - <html lang="...">
    - Skip link + <main id="main"> landmark
    - Ensure headings start with one <h1> (if missing, clone from <title>)
    - For tables in appendix: add <caption> if missing; mark header cells with scope=col
    """
    try:
        doc = _html.fromstring(html)
    except Exception:
        return html  # fail-soft

    _ensure_lang(doc, lang)
    _ensure_skiplink_and_main(doc)
    _ensure_h1_from_title(doc)
    _improve_tables(doc)

    return _html.tostring(doc, encoding="unicode", pretty_print=False, method="html")

# ---- helpers ----


def _ensure_lang(doc, lang: str):
    html_el = doc.xpath("//html")
    if html_el:
        el = html_el[0]
        if not el.get("lang"):
            el.set("lang", lang)


def _ensure_skiplink_and_main(doc):
    body = doc.xpath("//body")
    if not body:
        return
    body = body[0]

    # <main id="main"> wrapping content (if not present)
    mains = doc.xpath("//main")
    if not mains:
        # Wrap existing body children in <main>
        main = _html.Element("main")
        main.set("id", "main")
        # move all children
        for child in list(body):
            main.append(child)  # moves node
        body.append(main)

    # Skip link at top of body if not present
    has_skip = doc.xpath("//a[@class='skip-link' and @href='#main']")
    if not has_skip:
        a = _html.Element("a")
        a.set("href", "#main")
        a.set("class", "skip-link")
        a.text = "本文へスキップ"
        body.insert(0, a)
        # Basic inline style to keep visible for keyboard users but unobtrusive
        style = doc.xpath("//style")
        style_el = style[0] if style else None
        css = ".skip-link{position:absolute;left:-10000px;top:auto;width:1px;height:1px;overflow:hidden;} .skip-link:focus{position:static;width:auto;height:auto;}"
        if style_el is not None:
            style_el.text = (style_el.text or "") + "\n" + css
        else:
            head = doc.xpath("//head")
            if head:
                st = _html.Element("style")
                st.text = css
                head[0].append(st)


def _ensure_h1_from_title(doc):
    # If no <h1>, synthesize one from <title> and place at the top of <main>
    h1 = doc.xpath("//h1")
    if h1:
        return
    title_nodes = doc.xpath("//title")
    if not title_nodes or (title_nodes[0].text or "").strip() == "":
        return
    title_text = (title_nodes[0].text or "").strip()
    main = doc.xpath("//main")
    if not main:
        return
    h = _html.Element("h1")
    h.text = title_text
    # Insert at top of main
    main_el = main[0]
    main_el.insert(0, h)


def _improve_tables(doc):
    # Add caption for appendix table if missing; mark TH scope
    for table in doc.xpath("//section[contains(@class,'appendix')]//table"):
        # caption
        caps = table.xpath("./caption")
        if not caps:
            cap = _html.Element("caption")
            cap.text = "付録：その他のフィールド一覧"
            table.insert(0, cap)
        # th scope
        for th in table.xpath(".//thead//th"):
            if not th.get("scope"):
                th.set("scope", "col")
