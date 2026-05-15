#!/usr/bin/env python3
"""
extract_zotero_refs.py
======================

Extract Zotero citation metadata embedded inside a Microsoft Word (.docx)
document and emit a file that can be imported back into Zotero (or any
reference manager that reads CSL JSON / BibTeX).

The Zotero Word plug-in stores each in-text citation as a Word *field*
whose instruction text begins with the marker:

    ADDIN ZOTERO_ITEM CSL_CITATION { ...CSL JSON... }

This script:

  1. Opens the .docx as a ZIP archive.
  2. Reads every XML part that can contain body text (document, footnotes,
     endnotes, headers, footers).
  3. Walks the XML, joining `<w:instrText>` runs that belong to the same
     field, and pulls out the ADDIN ZOTERO_ITEM CSL_CITATION payload.
  4. Parses the embedded CSL JSON and harvests every `itemData` object.
  5. Deduplicates references (DOI first, then title+year+first-author).
  6. Writes the result as CSL JSON and/or BibTeX.

Usage
-----
    python extract_zotero_refs.py input.docx --output references.bib
    python extract_zotero_refs.py input.docx --output refs.json --format csljson
    python extract_zotero_refs.py input.docx --output refs.bib --also-json refs.json

Only the Python standard library is required.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
from typing import Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# WordprocessingML namespace — every w:* element lives here.
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# XML parts inside the .docx ZIP that can carry body text (and therefore
# Zotero fields). We try each one and silently skip any that aren't present.
DOCX_TEXT_PARTS = (
    "word/document.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    # Headers/footers/comments can in principle contain citations too.
    # They are matched dynamically below.
)

# Marker that identifies a Zotero in-text citation field.
ZOTERO_ITEM_MARKER = "ADDIN ZOTERO_ITEM CSL_CITATION"


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------

class ExtractionError(Exception):
    """Raised for any user-visible failure (bad file, no citations, etc.)."""


# ---------------------------------------------------------------------------
# Stage 1: load the relevant XML parts from the .docx
# ---------------------------------------------------------------------------

def load_docx_parts(docx_path: str) -> Dict[str, str]:
    """
    Open `docx_path` as a ZIP and return a dict mapping part name -> XML text
    for every part that can hold body content.
    """
    if not os.path.isfile(docx_path):
        raise ExtractionError(f"File not found: {docx_path}")

    try:
        zf = zipfile.ZipFile(docx_path)
    except zipfile.BadZipFile as exc:
        raise ExtractionError(
            f"{docx_path!r} is not a valid .docx (ZIP) file: {exc}"
        ) from exc

    parts: Dict[str, str] = {}
    with zf:
        names = set(zf.namelist())

        # Required: at least word/document.xml must exist.
        if "word/document.xml" not in names:
            raise ExtractionError(
                f"{docx_path!r} does not look like a Word document "
                "(missing word/document.xml)."
            )

        # Collect the named parts plus any header*/footer*/comments* XMLs.
        candidate_parts = set(DOCX_TEXT_PARTS)
        for n in names:
            if n.startswith("word/header") and n.endswith(".xml"):
                candidate_parts.add(n)
            elif n.startswith("word/footer") and n.endswith(".xml"):
                candidate_parts.add(n)
            elif n == "word/comments.xml":
                candidate_parts.add(n)

        for part in sorted(candidate_parts):
            if part in names:
                try:
                    parts[part] = zf.read(part).decode("utf-8")
                except UnicodeDecodeError:
                    # Word files should be UTF-8; if not, fall back leniently.
                    parts[part] = zf.read(part).decode("utf-8", errors="replace")

    return parts


# ---------------------------------------------------------------------------
# Stage 2: walk the XML and yield complete Zotero field-instruction strings
# ---------------------------------------------------------------------------

def iter_zotero_fields(xml_text: str) -> Iterable[str]:
    """
    Yield each Zotero ADDIN ZOTERO_ITEM CSL_CITATION instruction string found
    in `xml_text`.

    Word may split a long field instruction across several <w:instrText>
    elements. They are delimited by <w:fldChar w:fldCharType="begin"/> ...
    <w:fldChar w:fldCharType="end"/> (and there may be a "separate" in the
    middle, after which the run contains the *rendered* citation text rather
    than instruction text).

    We therefore parse the XML and track field nesting, concatenating
    instrText runs that fall inside a begin..separate/end pair.
    """
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise ExtractionError(f"Failed to parse Word XML: {exc}") from exc

    fld_char_tag = f"{{{W_NS}}}fldChar"
    instr_text_tag = f"{{{W_NS}}}instrText"
    type_attr = f"{{{W_NS}}}fldCharType"

    # We walk in document order. A stack handles nested fields (rare, but
    # legal). Each stack frame collects instrText pieces for that field.
    stack: List[List[str]] = []

    for elem in root.iter():
        tag = elem.tag
        if tag == fld_char_tag:
            ftype = elem.get(type_attr)
            if ftype == "begin":
                stack.append([])
            elif ftype in ("separate", "end"):
                # Field instruction ended (further runs are display text).
                if stack:
                    pieces = stack.pop()
                    instr = "".join(pieces).strip()
                    if ZOTERO_ITEM_MARKER in instr:
                        yield instr
                    # If "separate", any inner field still needs its own frame
                    # — handled naturally because we only pop the top one.
        elif tag == instr_text_tag and stack:
            # Accumulate this run into the currently-open field instruction.
            if elem.text:
                stack[-1].append(elem.text)


# ---------------------------------------------------------------------------
# Stage 3: pull the CSL JSON out of the field instruction
# ---------------------------------------------------------------------------

def parse_csl_citation(instr_text: str) -> List[dict]:
    """
    Given a Zotero field-instruction string, return a list of CSL JSON
    item dictionaries (the contents of each citationItems[i].itemData).
    """
    idx = instr_text.find(ZOTERO_ITEM_MARKER)
    if idx < 0:
        return []

    # Skip the marker; the JSON object begins at the next '{'.
    after_marker = instr_text[idx + len(ZOTERO_ITEM_MARKER):]
    brace_start = after_marker.find("{")
    if brace_start < 0:
        return []

    json_blob = _extract_balanced_json(after_marker, brace_start)
    if json_blob is None:
        return []

    try:
        data = json.loads(json_blob)
    except json.JSONDecodeError:
        # Zotero sometimes uses smart quotes inside titles; the JSON itself is
        # always ASCII-clean, so a real parse failure means the field is
        # malformed. Skip it rather than crashing the whole run.
        return []

    items: List[dict] = []
    for ci in data.get("citationItems", []) or []:
        item = ci.get("itemData")
        if isinstance(item, dict):
            items.append(item)
    return items


def _extract_balanced_json(s: str, start: int) -> Optional[str]:
    """
    Return the substring of `s` starting at `start` (which must be '{') that
    contains a balanced JSON object. Respects strings and escapes.
    """
    if start >= len(s) or s[start] != "{":
        return None

    depth = 0
    in_string = False
    escape = False
    for i in range(start, len(s)):
        c = s[i]
        if in_string:
            if escape:
                escape = False
            elif c == "\\":
                escape = True
            elif c == '"':
                in_string = False
        else:
            if c == '"':
                in_string = True
            elif c == "{":
                depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0:
                    return s[start:i + 1]
    return None


# ---------------------------------------------------------------------------
# Stage 4: deduplicate references
# ---------------------------------------------------------------------------

def _normalise_doi(doi: Optional[str]) -> Optional[str]:
    if not doi or not isinstance(doi, str):
        return None
    d = doi.strip().lower()
    # Strip common URL prefixes.
    for prefix in ("https://doi.org/", "http://doi.org/", "doi:"):
        if d.startswith(prefix):
            d = d[len(prefix):]
    return d or None


def _year_of(item: dict) -> Optional[str]:
    issued = item.get("issued") or {}
    parts = issued.get("date-parts") if isinstance(issued, dict) else None
    if isinstance(parts, list) and parts and isinstance(parts[0], list) and parts[0]:
        return str(parts[0][0])
    # Some Zotero exports use a flat 'literal' date.
    if isinstance(issued, dict) and isinstance(issued.get("literal"), str):
        m = re.search(r"\b(1[6-9]\d{2}|20\d{2}|21\d{2})\b", issued["literal"])
        if m:
            return m.group(1)
    return None


def _first_author_family(item: dict) -> Optional[str]:
    authors = item.get("author")
    if isinstance(authors, list) and authors:
        a = authors[0]
        if isinstance(a, dict):
            for key in ("family", "literal", "name"):
                v = a.get(key)
                if isinstance(v, str) and v.strip():
                    return v.strip()
    return None


def _dedup_key(item: dict) -> Tuple:
    doi = _normalise_doi(item.get("DOI"))
    if doi:
        return ("doi", doi)
    title = (item.get("title") or "").strip().lower()
    title = re.sub(r"\s+", " ", title)
    return ("tya", title, _year_of(item) or "", (_first_author_family(item) or "").lower())


def _completeness(item: dict) -> int:
    """Rough score: more populated fields = better. Used to pick a winner
    when two duplicates differ in detail."""
    return sum(1 for v in item.values() if v not in (None, "", [], {}))


def dedupe(items: List[dict]) -> List[dict]:
    """Collapse duplicates, preferring the metadata-richest variant."""
    best: Dict[Tuple, dict] = {}
    for it in items:
        key = _dedup_key(it)
        if key not in best or _completeness(it) > _completeness(best[key]):
            best[key] = it
    return list(best.values())


# ---------------------------------------------------------------------------
# Stage 5: exporters
# ---------------------------------------------------------------------------

def write_csl_json(items: List[dict], path: str) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(items, fh, ensure_ascii=False, indent=2)


# CSL itemType -> BibTeX entry type. Zotero uses the CSL vocabulary.
CSL_TO_BIBTEX_TYPE = {
    "article-journal": "article",
    "article-magazine": "article",
    "article-newspaper": "article",
    "article": "article",
    "paper-conference": "inproceedings",
    "chapter": "incollection",
    "book": "book",
    "thesis": "phdthesis",
    "report": "techreport",
    "manuscript": "unpublished",
    "webpage": "misc",
    "post-weblog": "misc",
    "post": "misc",
    "dataset": "misc",
    "software": "misc",
}


def _bibtex_escape(value: str) -> str:
    """Escape special BibTeX characters inside a {...} value."""
    if not isinstance(value, str):
        value = str(value)
    # Order matters: backslash first.
    replacements = {
        "\\": r"\textbackslash{}",
        "{": r"\{",
        "}": r"\}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "^": r"\^{}",
        "~": r"\~{}",
    }
    out = []
    for ch in value:
        out.append(replacements.get(ch, ch))
    return "".join(out)


def _format_authors_bibtex(authors: Optional[list]) -> Optional[str]:
    if not isinstance(authors, list) or not authors:
        return None
    formatted = []
    for a in authors:
        if not isinstance(a, dict):
            continue
        if "literal" in a and a["literal"]:
            # Organisational author — wrap in braces so BibTeX won't reorder.
            formatted.append("{" + str(a["literal"]) + "}")
            continue
        family = (a.get("family") or "").strip()
        given = (a.get("given") or "").strip()
        if family and given:
            formatted.append(f"{family}, {given}")
        elif family:
            formatted.append(family)
        elif given:
            formatted.append(given)
    return " and ".join(formatted) if formatted else None


def _safe_citekey(item: dict, used: set) -> str:
    """Build a deterministic citation key: AuthorYearWord."""
    author = _first_author_family(item) or "anon"
    author = re.sub(r"[^A-Za-z]", "", author) or "anon"
    year = _year_of(item) or "nd"
    title = item.get("title") or ""
    # First significant word of the title (>=4 letters), lowercased.
    word = ""
    for w in re.findall(r"[A-Za-z]+", title):
        if len(w) >= 4:
            word = w.lower()
            break
    base = f"{author}{year}{word}".strip() or "ref"
    key = base
    n = 1
    while key in used:
        n += 1
        key = f"{base}{n}"
    used.add(key)
    return key


def _csl_item_to_bibtex_fields(item: dict) -> List[Tuple[str, str]]:
    """Map a CSL JSON item to an ordered list of (bibtex_field, value)."""
    out: List[Tuple[str, str]] = []

    title = item.get("title")
    if title:
        out.append(("title", str(title)))

    authors = _format_authors_bibtex(item.get("author"))
    if authors:
        out.append(("author", authors))

    editors = _format_authors_bibtex(item.get("editor"))
    if editors:
        out.append(("editor", editors))

    year = _year_of(item)
    if year:
        out.append(("year", year))

    container = item.get("container-title")
    csl_type = item.get("type", "")
    if container:
        if csl_type in ("paper-conference",):
            out.append(("booktitle", str(container)))
        elif csl_type in ("chapter",):
            out.append(("booktitle", str(container)))
        else:
            out.append(("journal", str(container)))

    for csl_field, bib_field in (
        ("volume", "volume"),
        ("issue", "number"),
        ("page", "pages"),
        ("publisher", "publisher"),
        ("publisher-place", "address"),
        ("DOI", "doi"),
        ("URL", "url"),
        ("ISBN", "isbn"),
        ("ISSN", "issn"),
        ("abstract", "abstract"),
    ):
        v = item.get(csl_field)
        if v:
            out.append((bib_field, str(v)))

    return out


def write_bibtex(items: List[dict], path: str) -> None:
    used_keys: set = set()
    with open(path, "w", encoding="utf-8") as fh:
        for item in items:
            csl_type = item.get("type") or ""
            bib_type = CSL_TO_BIBTEX_TYPE.get(csl_type, "misc")
            key = _safe_citekey(item, used_keys)
            fh.write(f"@{bib_type}{{{key},\n")
            for field, value in _csl_item_to_bibtex_fields(item):
                fh.write(f"  {field} = {{{_bibtex_escape(value)}}},\n")
            fh.write("}\n\n")


# ---------------------------------------------------------------------------
# Top-level pipeline + CLI
# ---------------------------------------------------------------------------

def extract_references(docx_path: str) -> List[dict]:
    """Run the full pipeline and return the deduplicated list of CSL items."""
    parts = load_docx_parts(docx_path)

    raw_items: List[dict] = []
    fields_seen = 0
    for _name, xml in parts.items():
        for instr in iter_zotero_fields(xml):
            fields_seen += 1
            raw_items.extend(parse_csl_citation(instr))

    if fields_seen == 0:
        raise ExtractionError(
            "No Zotero citation fields found. The document may use plain-text "
            "citations, a different citation tool (e.g. Mendeley/EndNote), or "
            "the citations may have been unlinked."
        )

    if not raw_items:
        raise ExtractionError(
            "Zotero fields were present but contained no usable CSL JSON. "
            "The document may have been saved by a very old Zotero plug-in."
        )

    return dedupe(raw_items)


def _infer_format(path: str, explicit: Optional[str]) -> str:
    if explicit:
        return explicit.lower()
    ext = os.path.splitext(path)[1].lower()
    if ext in (".bib", ".bibtex"):
        return "bibtex"
    if ext in (".json",):
        return "csljson"
    raise ExtractionError(
        f"Cannot infer output format from extension {ext!r}. "
        "Pass --format bibtex or --format csljson."
    )


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(
        description="Extract Zotero references from a .docx file."
    )
    p.add_argument("input", help="Path to the .docx file.")
    p.add_argument(
        "-o", "--output",
        required=True,
        help="Output file path (.bib for BibTeX, .json for CSL JSON).",
    )
    p.add_argument(
        "-f", "--format",
        choices=("bibtex", "csljson"),
        help="Force output format; default is inferred from --output extension.",
    )
    p.add_argument(
        "--also-json",
        metavar="PATH",
        help="Additionally write a CSL JSON copy to PATH.",
    )
    p.add_argument(
        "-v", "--verbose", action="store_true",
        help="Print a one-line summary per reference found.",
    )

    args = p.parse_args(argv)

    try:
        items = extract_references(args.input)
        fmt = _infer_format(args.output, args.format)

        if fmt == "bibtex":
            write_bibtex(items, args.output)
        elif fmt == "csljson":
            write_csl_json(items, args.output)
        else:
            raise ExtractionError(f"Unknown output format: {fmt}")

        if args.also_json and fmt != "csljson":
            write_csl_json(items, args.also_json)

    except ExtractionError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    print(f"Extracted {len(items)} unique reference(s) → {args.output}")
    if args.also_json:
        print(f"CSL JSON copy → {args.also_json}")
    if args.verbose:
        for it in items:
            title = (it.get("title") or "").strip()
            year = _year_of(it) or "n.d."
            author = _first_author_family(it) or "?"
            print(f"  - {author} ({year}) {title[:80]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
