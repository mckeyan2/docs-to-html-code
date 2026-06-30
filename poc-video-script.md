# POC Video Narration Script
## Doc-to-HTML Converter — Proof of Concept

---

### INTRO

[VISUAL: poc.html page loading in browser — hero section visible]

"Welcome to the Doc-to-HTML Converter — a browser-based tool that transforms Word documents, spreadsheets, and plain text into clean, semantic HTML5 — all without uploading your files to any server.

This is the proof-of-concept walkthrough. Let's go through exactly what this tool does, how it works, and why every output it produces is production-ready."

---

### HOW IT WORKS — 4 STEPS

[VISUAL: Scroll to the "How It Works" section — four numbered step cards]

"The conversion happens in four steps.

**Step one** — you upload your file. Drag and drop a Word document, an Excel spreadsheet, a CSV, or just paste plain text into the editor. Nothing leaves your machine.

**Step two** — the tool extracts the content. For `.docx` files, we use Mammoth.js to pull structured text out of the Word XML. For `.xlsx`, `.xls`, and `.csv`, SheetJS parses the spreadsheet into rows and columns.

**Step three** — the content passes through the conversion pipeline. Fifteen production rules run in sequence, normalising headings, converting phone numbers to clickable links, replacing symbol characters with HTML entities, stripping legacy Microsoft Office markup, and more.

**Step four** — you get clean HTML5 output. Preview it in the result panel, copy the markup, or download it as a `.txt` file ready to paste into your CMS."

---

### SAMPLE CONVERSION — SIDE BY SIDE

[VISUAL: Scroll to "Sample Conversion" section — left box (Word-style doc) and right box (HTML output) visible side by side]

"Here's a concrete before-and-after.

On the left is a Word-style document — the kind of content you'd typically paste from an email or a client brief. It has a heading, a phone number, a sentence with a trademark symbol, a bold term, an underlined word, and a data table.

On the right is what the converter produces.

[VISUAL: Point to heading row]
The top-level heading has been normalised from H1 down to H2 — because H1 belongs to the page, not the content block.

[VISUAL: Point to phone number line]
The phone number `044-1234567` has become a proper `tel:` anchor — tap it on mobile and it dials directly.

[VISUAL: Point to trademark line]
The trademark symbol in the text has been converted to the `&trade;` HTML entity — correct, portable, and accessible.

[VISUAL: Point to bold and underline lines]
The bold word is now wrapped in a semantic `<strong>` tag. The underline — which means nothing in HTML — has been stripped, leaving clean plain text.

[VISUAL: Point to the table]
And the table has been given proper structure: `<thead>` with `<th scope='col'>` for the header row, and `<tbody>` for the data rows. Screen readers will navigate it correctly."

---

### CONVERSION RULES TABLE

[VISUAL: Scroll to the "Conversion Rules" section — full table visible]

"These are the fifteen production rules that run on every conversion.

**Rule 1** — Heading normalisation. Any H1 in the pasted content is shifted to H2 to preserve the page's semantic hierarchy.

**Rule 2** — Phone numbers become `tel:` links. The rule detects common Australian phone formats, strips spaces and dashes, and wraps the number in an anchor. Fax numbers are identified by a 'Fax:' label and skipped.

**Rule 3** — Superscript and subscript characters are converted to Unicode. So `x` superscript `2` becomes `x²`, and H subscript 2 O becomes `H₂O` — no extra tags needed in the output.

**Rule 4** — Bold tags are semanticised. Legacy `<b>` tags become `<strong>`.

**Rule 5** — Italic handling. Short italic phrases — less than three words — are converted to `<strong>` for emphasis. Longer italic spans are unwrapped entirely, since they're usually formatting noise.

**Rule 6** — Underline removal. HTML underline has no semantic meaning; the `<u>` tag is stripped and the inner text is preserved.

**Rule 7** — XML and MSO comments are stripped. Word documents are full of Microsoft Office processing instructions and XML namespaces. This rule removes all of them.

**Rule 8** — Symbol entities. Copyright, trademark, and registered trademark symbols — whether typed directly or written as bracketed text — are replaced with `&copy;`, `&trade;`, and `&reg;`.

**Rule 9** — Table scope. When the first row of a table contains bold text, that row is promoted to a `<thead>` with `<th scope='col'>` cells.

**Rule 10** — Bluecard® auto-append. When the word 'Bluecard' appears without the registered mark, the `®` symbol is appended automatically.

**Rule 11** — Non-breaking spaces are replaced with regular spaces.

**Rule D** — Strikethrough text is removed. `<s>`, `<del>`, and Markdown-style `~~tildes~~` are all deleted — deleted text shouldn't carry forward into published HTML.

**Rule E** — Strong-wraps-anchor correction. When a `<strong>` tag is nested inside an `<a>` tag, the order is swapped so that `<strong>` wraps `<a>`. This matches the HTML5 content model.

**Rule F** — File URL correction. Any `file://` URL — an absolute local path that makes no sense on the web — is rewritten to `https://`.

**Rule A** — No bold inside headings. Headings are already visually prominent; a `<strong>` tag inside an `<h2>` or `<h3>` is redundant. The rule unwraps it.

**Rule C** — Consecutive bold tags are merged. Two adjacent `<strong>` blocks with the same text content are collapsed into one."

---

### POST-CONVERSION TEST REPORT

[VISUAL: Scroll to the "Post-Conversion Test Report" section — scorecard tiles at top, then checklist, then feature cards]

"Every conversion rule has automated tests. Here's the test report.

[VISUAL: Scorecard — 47 total, 47 passed, 0 failed, 100%]

Forty-seven tests. Forty-seven passed. Zero failed. One hundred percent coverage.

[VISUAL: Scroll through the checklist]

The checklist maps every test group to its production rule. Heading normalisation — checked. Phone links — checked. Unicode sup/sub — checked. Table scope headers — checked. All fifteen rules verified.

[VISUAL: Feature cards below checklist]

Below the checklist are the test suite capabilities: in-app live execution so tests run every time the page loads, a one-click download to export the full report as a `.txt` file, and expandable test cards so you can inspect each individual assertion inline."

---

### STANDALONE TEST SUITE

[VISUAL: Scroll to the "Standalone Test Suite" section — three info cards]

"Alongside the main app, there's a dedicated `tests.html` page.

Open it in any browser, no build step, no server. The forty-seven tests run automatically on page load. Each test group is collapsible. Pass or fail is visible at a glance. And the download report button exports everything for sharing with stakeholders or attaching to a pull request."

---

### KEY FEATURES

[VISUAL: Scroll to the "Key Features" section — six feature cards]

"Six things that make this tool practical for real content workflows.

**Browser-only processing** — your documents never leave the device. There's no upload, no cloud API, no data retention risk.

**Multi-format support** — `.docx`, `.xlsx`, `.xls`, `.csv`, and direct paste from clipboard or editor.

**Semantic HTML5 output** — every element has the right tag. Headings, tables, links, and emphasis are all structurally correct, not just visually styled.

**47 automated tests** — the conversion logic is fully test-covered. Regressions are caught before they reach production.

**One-click copy and download** — copy the HTML to the clipboard or download it as a `.txt` file. Paste directly into any CMS, email builder, or code editor.

**Rich text editor** — if you don't have a file, type or paste into the built-in editor. Full toolbar: headings, bold, italic, lists, alignment, links, and more."

---

### OUTRO

[VISUAL: Scroll back to the top — hero section]

"That's the Doc-to-HTML Converter.

Clean output. Fifteen production rules. Forty-seven tests. No server. No upload.

Drop a file, get HTML."

---

*Estimated read time: 4–5 minutes at a natural speaking pace.*
