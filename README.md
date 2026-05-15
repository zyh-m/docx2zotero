# docx2zotero

Recover Zotero references from a Microsoft Word `.docx` file when the cited items are no longer in your local Zotero library.

If a collaborator hands you a Word document with Zotero-formatted citations and you don't have the underlying references, `docx2zotero` digs the embedded CSL JSON metadata out of the document's field codes, deduplicates it, and writes a BibTeX (or CSL JSON) file you can import straight into Zotero.

Pure Python 3.8+, **standard library only**, single file.

## Why this exists

Zotero's Word plug-in embeds full bibliographic metadata inside every in-text citation as a Word field — not just a key or a hyperlink. That metadata stays in the file even on a machine where Zotero has never been installed. As long as the citations haven't been "unlinked," they can be fully recovered. This tool does that recovery.

## Install

```bash
git clone https://github.com/<you>/docx2zotero.git
cd docx2zotero
```

No dependencies. If you'd like the script on your `$PATH`:

```bash
chmod +x docx2zotero.py
ln -s "$(pwd)/docx2zotero.py" ~/.local/bin/docx2zotero
```

## Usage

```bash
python docx2zotero.py paper.docx --output refs.bib
```

Common options:

```bash
# Force a specific output format (otherwise inferred from the extension)
python docx2zotero.py paper.docx -o refs.json -f csljson

# Write BibTeX and a CSL JSON copy in one go
python docx2zotero.py paper.docx -o refs.bib --also-json refs.json

# Print one summary line per recovered reference
python docx2zotero.py paper.docx -o refs.bib --verbose
```

Exit codes: `0` on success, `2` on any user-visible error (bad file, no Zotero fields found, etc.) with a clear message on stderr.

## Importing into Zotero

The cleanest way to land everything inside a specific collection:

1. In Zotero, create (or pick) your target collection — e.g. `Recovered — Paper X`.
2. **Right-click the collection itself** → **Import…** → **A file (BibTeX, RIS, Zotero RDF…)** → choose `refs.bib`.
3. Leave **"Place imported collection into selected collection"** unticked. Click **Import**.

All recovered references land inside that collection. To catch any items that overlap with what you already have, use **My Library → Duplicate Items** afterwards and merge.

> Importing via **File → Import…** from the top menu sometimes drops items into the library root instead of the selected collection. Right-clicking the collection is the reliable path.

## How it works

A `.docx` is a ZIP archive of XML files. When you insert a citation via Zotero, the plug-in writes a Word field whose instruction text looks like this:

```
ADDIN ZOTERO_ITEM CSL_CITATION { ...JSON... }
```

The JSON is a [CSL JSON citation](https://github.com/citation-style-language/schemas/blob/master/csl-citation.json) whose `citationItems[i].itemData` carries the full bibliographic record (title, authors, year, DOI, container-title, pages, etc.).

The pipeline:

1. **Unzip** the document and read every XML part that can hold body text (`word/document.xml`, `word/footnotes.xml`, `word/endnotes.xml`, plus headers/footers/comments when present).
2. **Walk** the XML tracking `<w:fldChar w:fldCharType="begin"/>` … `end` pairs, concatenating any `<w:instrText>` runs that fall inside a single field (Word can split long instruction text across several runs).
3. **Parse** each field's instruction. Strip the `ADDIN ZOTERO_ITEM CSL_CITATION` marker, then read a balanced JSON object out of what follows.
4. **Deduplicate**. Primary key: normalised DOI (case- and prefix-folded — `https://doi.org/`, `doi:`, etc. all collapse). Fallback key: `(lowercased-title, year, first-author-family)`. When duplicates collide, the metadata-richer variant wins.
5. **Export** as CSL JSON (lossless) and/or BibTeX. The BibTeX writer maps CSL item types (`article-journal` → `@article`, `paper-conference` → `@inproceedings`, etc.), escapes BibTeX-special characters, and generates deterministic `AuthorYearWord` citekeys.

## Limitations

| Limitation | What you'll see | Recourse |
|---|---|---|
| Citations have been **unlinked** in Word | "No Zotero citation fields found." | Unrecoverable from the document alone — the field metadata is gone. Re-look-up titles via Crossref / Google Scholar. |
| Document uses **Mendeley** or **EndNote** instead of Zotero | "No Zotero citation fields found." | Those tools use different `ADDIN` markers (`ADDIN CSL_CITATION`, `ADDIN EN.CITE`). Extending support is a few lines in `iter_zotero_fields` / `parse_csl_citation` — PRs welcome. |
| Reference is in the bibliography only, never cited in-text | Missing from the output | This is by design of the Zotero plug-in: full per-item metadata is only stored in the `ZOTERO_ITEM` (in-text) fields. The `ZOTERO_BIBL` block at the end just references item URIs. |
| Pre-CSL legacy Zotero plug-in versions | Marker not detected | Not supported. Open an issue with a sample file. |
| CSL → BibTeX is lossy | Some CSL fields don't map (e.g. `event-place`, `archive_location`) | Use `--also-json` to keep a lossless CSL JSON copy alongside the .bib. |

## Output formats

- **CSL JSON** — Zotero's native interchange format. Lossless. Recommended if you only need to repopulate Zotero.
- **BibTeX** — for LaTeX workflows and the widest tool compatibility. Slightly lossy.

RIS export is not currently implemented; CSL JSON covers the same use case more cleanly. Add `--format ris` support by writing a `write_ris` exporter alongside `write_bibtex` if needed.

## Project layout

```
docx2zotero/
├── docx2zotero.py     # the entire tool — single file, stdlib only
├── README.md
└── LICENSE
```

## Development

A quick sanity test without any real Zotero documents handy:

```python
import json, zipfile

item = {
    "id": 1, "type": "article-journal",
    "title": "Attention Is All You Need",
    "author": [{"family": "Vaswani", "given": "Ashish"}],
    "issued": {"date-parts": [[2017]]},
    "DOI": "10.48550/arXiv.1706.03762",
}
payload = json.dumps({"citationItems": [{"id": 1, "itemData": item}]})
doc = f'''<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve">ADDIN ZOTERO_ITEM CSL_CITATION {payload}</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p></w:body></w:document>'''
ct = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
with zipfile.ZipFile("sample.docx", "w") as zf:
    zf.writestr("[Content_Types].xml", ct)
    zf.writestr("_rels/.rels", rels)
    zf.writestr("word/document.xml", doc)
```

Then:

```bash
python docx2zotero.py sample.docx -o out.bib --verbose
```

## License

MIT. See `LICENSE`.
