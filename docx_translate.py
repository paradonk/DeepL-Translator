"""Shared DeepL translation engine for .docx files.

Works directly on the document XML, so it:
  * keeps run formatting (bold, italic, fonts, colours, hyperlinks) by sending
    each paragraph to DeepL as tagged XML and writing the result back run by run;
  * translates the body, tables, text boxes, every header/footer, footnotes and
    endnotes, on any OS, without Microsoft Word;
  * sends texts in batches and translates repeated paragraphs only once.
"""
from __future__ import annotations

import copy
import os
import re
import zipfile
from pathlib import Path
from typing import BinaryIO, Callable, Iterable, Optional, Union
from xml.sax.saxutils import escape

import deepl
from dotenv import load_dotenv
from lxml import etree

APP_DIR = Path(__file__).resolve().parent
ENV_PATH = APP_DIR / ".env"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W_P = f"{{{W_NS}}}p"
W_T = f"{{{W_NS}}}t"
MC_FALLBACK = f"{{{MC_NS}}}Fallback"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Body, all headers/footers, footnotes and endnotes. Text boxes live inside these.
TRANSLATABLE_PART = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")

# DeepL allows 50 texts per request and a 128 KiB request body.
MAX_BATCH_TEXTS = 50
MAX_BATCH_CHARS = 60_000

# Fallback list, used when the live list can't be fetched. Bare "EN" and "PT"
# are deliberately absent: DeepL wants EN-GB/EN-US and PT-PT/PT-BR.
TARGET_LANGUAGES: list[tuple[str, str]] = [
    ("AR", "Arabic"), ("BG", "Bulgarian"), ("CS", "Czech"), ("DA", "Danish"),
    ("DE", "German"), ("EL", "Greek"), ("EN-GB", "English (British)"),
    ("EN-US", "English (American)"), ("ES", "Spanish"), ("ET", "Estonian"),
    ("FI", "Finnish"), ("FR", "French"), ("HE", "Hebrew"), ("HU", "Hungarian"),
    ("ID", "Indonesian"), ("IT", "Italian"), ("JA", "Japanese"), ("KO", "Korean"),
    ("LT", "Lithuanian"), ("LV", "Latvian"), ("NB", "Norwegian (Bokmål)"),
    ("NL", "Dutch"), ("PL", "Polish"), ("PT-BR", "Portuguese (Brazilian)"),
    ("PT-PT", "Portuguese (European)"), ("RO", "Romanian"), ("RU", "Russian"),
    ("SK", "Slovak"), ("SL", "Slovenian"), ("SV", "Swedish"), ("TH", "Thai"),
    ("TR", "Turkish"), ("UK", "Ukrainian"), ("VI", "Vietnamese"),
    ("ZH-HANS", "Chinese (Simplified)"), ("ZH-HANT", "Chinese (Traditional)"),
]
_AMBIGUOUS_CODES = {"EN": "EN-GB", "PT": "PT-PT", "ZH": "ZH-HANS"}

ProgressFn = Callable[[int, int], None]


def load_api_key() -> str:
    """Read DEEPL_API_KEY from the environment or the .env next to this file."""
    load_dotenv(ENV_PATH)
    return os.getenv("DEEPL_API_KEY", "").strip()


def make_client(api_key: str) -> deepl.Translator:
    client_cls = getattr(deepl, "DeepLClient", deepl.Translator)
    return client_cls(api_key)


def normalize_lang(code: str) -> str:
    code = code.strip().upper()
    return _AMBIGUOUS_CODES.get(code, code)


def fetch_target_languages(client: deepl.Translator) -> list[tuple[str, str]]:
    """Live list from DeepL, without the ambiguous EN/PT/ZH codes."""
    langs = [(l.code.upper(), l.name) for l in client.get_target_languages()]
    langs = [(c, n) for c, n in langs if c not in _AMBIGUOUS_CODES]
    return sorted(langs, key=lambda cn: cn[1])


def needs_translation(text: str) -> bool:
    """Skip empty text and text with no letters (numbers, dates, symbols)."""
    return any(ch.isalpha() for ch in text)


# ---------------------------------------------------------------- XML helpers

# Line breaks and tabs split a paragraph into separately translated units.
_BREAKS = {f"{{{W_NS}}}{name}" for name in ("br", "cr", "tab", "ptab")}
# Other elements that sit between pieces of text and must stay between them.
_SEPARATORS = {f"{{{W_NS}}}{name}" for name in (
    "sym", "noBreakHyphen", "softHyphen", "drawing",
    "pict", "object", "fldChar", "footnoteReference", "endnoteReference")}
# Run properties that don't change how text looks (spell-check language etc.).
_IGNORED_RPR = {f"{{{W_NS}}}{name}" for name in ("lang", "noProof", "rPrChange")}


def _own_elements(p):
    """Descendants of paragraph p in document order, skipping paragraphs
    nested inside it (text boxes anchored in p) and legacy fallback copies."""
    stack = [iter(p)]
    while stack:
        for el in stack[-1]:
            if el.tag in (W_P, MC_FALLBACK):
                continue
            yield el
            stack.append(iter(el))
            break
        else:
            stack.pop()


def _format_key(t, p) -> tuple:
    """Two w:t nodes with equal keys look the same and sit in the same container
    (paragraph, hyperlink, tracked insertion...)."""
    run = t.getparent()
    rpr = run.find(f"{{{W_NS}}}rPr")
    rpr_bytes = b""
    if rpr is not None:
        rpr = etree.fromstring(etree.tostring(rpr))
        for child in list(rpr):
            if child.tag in _IGNORED_RPR:
                rpr.remove(child)
        rpr_bytes = etree.tostring(rpr)
    container = run.getparent()
    return (rpr_bytes, None if container is p else id(container))


class _Group:
    """Adjacent w:t nodes that share formatting and have nothing between them."""
    __slots__ = ("nodes", "key", "tagged", "silent")

    def __init__(self, node, key):
        self.nodes, self.key = [node], key
        self.tagged = True
        self.silent = False  # whitespace only: sent as plain text, emptied after

    @property
    def text(self) -> str:
        return "".join(t.text for t in self.nodes)

    def set_text(self, text: str) -> None:
        for i, t in enumerate(self.nodes):
            t.text = text if i == 0 else ""
            t.set(XML_SPACE, "preserve")


def _split_ws(text: str) -> tuple[str, str, str]:
    core = text.strip()
    if not core:
        return text, "", ""
    lead = text[: len(text) - len(text.lstrip())]
    return lead, core, text[len(lead) + len(core):]


class _Paragraph:
    """One paragraph prepared for DeepL.

    Text in the paragraph's dominant formatting is sent as plain text; only runs
    that look different (bold word, link, other font...) are wrapped in <s>
    tags. Keeping most of the sentence untagged lets DeepL translate it as a
    whole, so grammar agreement across formatting boundaries stays correct.
    """
    __slots__ = ("p", "groups", "source")

    def __init__(self, p, groups: list[_Group]):
        self.p = p
        self.groups = groups
        out = []
        for g in groups:
            if g.tagged and not g.silent:
                lead, core, trail = _split_ws(g.text)
                out.append(escape(lead))
                if core:
                    out.append(f"<s>{escape(core)}</s>")
                out.append(escape(trail))
            else:
                out.append(escape(g.text))
        self.source = "".join(out)

    @classmethod
    def from_p(cls, p) -> list["_Paragraph"]:
        """Split paragraph p at line breaks and tabs into translation units, so
        text can never be moved across a break."""
        units: list[list[_Group]] = [[]]
        separated = True
        for el in _own_elements(p):
            if el.tag in _BREAKS:
                if units[-1]:
                    units.append([])
                separated = True
            elif el.tag in _SEPARATORS:
                separated = True
            elif el.tag == W_T and el.text:
                key = _format_key(el, p)
                groups = units[-1]
                if not separated and groups and groups[-1].key == key:
                    groups[-1].nodes.append(el)
                else:
                    groups.append(_Group(el, key))
                separated = False
        return [cls._prepare(p, groups) for groups in units
                if groups and needs_translation("".join(g.text for g in groups))]

    @classmethod
    def _prepare(cls, p, groups: list[_Group]) -> "_Paragraph":
        # The dominant format goes untagged, except that two untagged groups
        # may not touch (e.g. both sides of a footnote mark), otherwise the
        # text between them couldn't be split back.
        weight: dict = {}
        for g in groups:
            weight[g.key] = weight.get(g.key, 0) + len(g.text.strip())
        base = max(weight, key=weight.get)
        # Prefer the least-formatted text (usually the "normal" words) when it
        # is a fair share of the paragraph, so tags mark the special runs.
        direct = [key for key in weight if key[1] is None]
        if direct:
            simplest = min(direct, key=lambda key: len(key[0]))
            if weight[simplest] * 3 >= sum(weight.values()):
                base = simplest
        prev_plain = False
        for g in groups:
            if not g.text.strip():
                # Whitespace-only: no tag of its own; its spaces travel with the
                # neighbouring text.
                g.silent = True
                continue
            g.tagged = not (g.key == base and not prev_plain)
            prev_plain = not g.tagged
        return cls(p, groups)

    def _top(self, node):
        """The ancestor of w:t node that is a direct child of the paragraph."""
        el = node.getparent()
        while el.getparent() is not self.p:
            el = el.getparent()
        return el

    def _new_plain_group(self, template: _Group, tagged: list[_Group], j: int) -> _Group:
        """A new run with the untagged formatting, placed between tagged[j-1]
        and tagged[j], for words DeepL moved to where no plain run existed."""
        run = etree.Element(f"{{{W_NS}}}r")
        rpr = template.nodes[0].getparent().find(f"{{{W_NS}}}rPr")
        if rpr is not None:
            run.append(copy.deepcopy(rpr))
        t = etree.SubElement(run, W_T)
        if j > 0:
            self._top(tagged[j - 1].nodes[-1]).addnext(run)
        else:
            self._top(tagged[0].nodes[0]).addprevious(run)
        return _Group(t, template.key)

    def apply(self, translated_xml: str) -> None:
        tagged = [g for g in self.groups if g.tagged and not g.silent]
        plain = [g for g in self.groups if not g.tagged]
        # gaps[j] = the group that receives the untagged text just before
        # tagged[j]; gaps[len(tagged)] = the one after the last tag. A plain
        # group wins over a whitespace-only one.
        gaps: list[Optional[_Group]] = [None] * (len(tagged) + 1)
        j = 0
        for g in self.groups:
            if g.tagged and not g.silent:
                j += 1
            elif gaps[j] is None or gaps[j].silent:
                gaps[j] = g

        try:
            root = etree.fromstring(f"<x>{translated_xml}</x>".encode("utf-8"))
            gap_texts = [root.text or ""] + [el.tail or "" for el in root]
            tag_texts = ["".join(el.itertext()) for el in root]
        except etree.XMLSyntaxError:
            text = re.sub(r"</?s>", "", translated_xml)
            text = text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
            gap_texts, tag_texts = [text], []

        k = len(tagged)
        if len(tag_texts) > k:  # DeepL split a tag: fold extras into the last one
            if k == 0:
                gap_texts = ["".join(gap_texts[i] + (tag_texts[i] if i < len(tag_texts) else "")
                                     for i in range(len(gap_texts)))]
                tag_texts = []
            else:
                extra = "".join(gap_texts[i] + tag_texts[i] for i in range(k, len(tag_texts)))
                tag_texts = tag_texts[: k - 1] + [tag_texts[k - 1] + gap_texts[k] + extra]
                gap_texts = gap_texts[:k] + [gap_texts[-1]]
        tag_texts += [""] * (k - len(tag_texts))
        gap_texts += [""] * (k + 1 - len(gap_texts))

        texts = {id(g): t for g, t in zip(tagged, tag_texts)}
        for j, gap_text in enumerate(gap_texts):
            holder = gaps[j]
            wrong_format = holder is not None and holder.silent and plain and holder.key != plain[0].key
            if (holder is None or wrong_format) and gap_text.strip() and plain and k:
                holder = self._new_plain_group(plain[0], tagged, j)
                self.groups.append(holder)
            if holder is not None:
                texts[id(holder)] = gap_text
            elif j > 0:   # only whitespace (or nowhere better): keep with neighbour
                texts[id(tagged[j - 1])] += gap_text
            elif k:
                texts[id(tagged[0])] = gap_text + texts[id(tagged[0])]
        for g in self.groups:
            g.set_text(texts.get(id(g), ""))


def _batches(texts: list[str]) -> Iterable[list[str]]:
    batch, size = [], 0
    for text in texts:
        if batch and (len(batch) >= MAX_BATCH_TEXTS or size + len(text) > MAX_BATCH_CHARS):
            yield batch
            batch, size = [], 0
        batch.append(text)
        size += len(text)
    if batch:
        yield batch


# ---------------------------------------------------------------- public API

def translate_docx(
    src: Union[str, Path, BinaryIO],
    dst: Union[str, Path, BinaryIO],
    target_lang: str,
    client: deepl.Translator,
    progress: Optional[ProgressFn] = None,
    log: Optional[Callable[[str], None]] = None,
) -> dict:
    """Translate src into dst. src/dst may be paths or binary file objects.

    progress(done, total) is called after each batch, counting unique text blocks.
    Returns {"paragraphs": unique text blocks sent, "characters": chars sent}.
    """
    target_lang = normalize_lang(target_lang)
    log = log or (lambda _msg: None)
    parser = etree.XMLParser(huge_tree=True, remove_blank_text=False)

    with zipfile.ZipFile(src) as zin:
        trees: dict[str, etree._Element] = {}
        paragraphs: list[_Paragraph] = []
        for name in zin.namelist():
            if not TRANSLATABLE_PART.match(name):
                continue
            root = etree.fromstring(zin.read(name), parser)
            trees[name] = root
            found = 0
            # Includes the legacy (mc:Fallback) copies of text boxes; they repeat
            # the main copy's text, so de-duplication makes them free.
            for p in root.iter(W_P):
                units = _Paragraph.from_p(p)
                paragraphs.extend(units)
                found += bool(units)
            if found:
                log(f"{name}: {found} paragraphs with text")

        unique = list(dict.fromkeys(p.source for p in paragraphs))
        total = len(unique)
        log(f"{len(paragraphs)} text blocks to translate ({total} unique)")
        if progress:
            progress(0, total)

        translated: dict[str, str] = {}
        for batch in _batches(unique):
            results = client.translate_text(batch, target_lang=target_lang, tag_handling="xml")
            translated.update(zip(batch, (r.text for r in results)))
            if progress:
                progress(len(translated), total)

        for p in paragraphs:
            p.apply(translated[p.source])

        # Rewrite the package entry by entry, keeping order and compression.
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename in trees:
                    data = etree.tostring(trees[info.filename], xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
                else:
                    data = zin.read(info.filename)
                zout.writestr(info, data)

    return {"paragraphs": total, "characters": sum(len(s) for s in unique)}


def output_path_for(src: Path, target_lang: str) -> Path:
    """report.docx -> report_fr_1729705432.docx, next to the source file."""
    import time
    code = normalize_lang(target_lang).lower()
    return src.with_name(f"{src.stem}_{code}_{int(time.time())}{src.suffix}")
