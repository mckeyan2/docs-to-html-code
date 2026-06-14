/* ─────────────────────────────────────────────────────────────────
   app.js  —  DOCX / XLSX / XLS / CSV / Text-Editor → HTML
   Requires: mammoth.min.js + xlsx.min.js
   ───────────────────────────────────────────────────────────────── */
(function () {
  'use strict';

  /* ══ DOM refs ══════════════════════════════════════════════════ */
  var dropZone     = document.getElementById('drop-zone');
  var fileInput    = document.getElementById('file-input');
  var uploadStatus = document.getElementById('upload-status');
  var editorStatus = document.getElementById('editor-status');
  var resultPanel  = document.getElementById('result-panel');
  var fileLabel    = document.getElementById('file-name-label');
  var previewPane  = document.getElementById('preview-pane');
  var codePane     = document.getElementById('code-pane');
  var dlBtn        = document.getElementById('dl-btn');
  var copyBtn      = document.getElementById('copy-btn');
  var resetBtn     = document.getElementById('reset-btn');
  var tabs         = document.querySelectorAll('.tab');
  var toast        = document.getElementById('toast');

  /* Editor refs */
  var editorBody       = document.getElementById('editor-body');
  var editorConvertBtn = document.getElementById('editor-convert-btn');
  var blockFormat      = document.getElementById('block-format');
  var tbClear          = document.getElementById('tb-clear');
  var tbLink           = document.getElementById('tb-link');
  var tbHr             = document.getElementById('tb-hr');

  /* Mode refs */
  var modeUploadBtn = document.getElementById('mode-upload-btn');
  var modeEditorBtn = document.getElementById('mode-editor-btn');
  var modeUpload    = document.getElementById('mode-upload');
  var modeEditor    = document.getElementById('mode-editor');

  /* ══ State ════════════════════════════════════════════════════ */
  var state = { htmlSource: '', baseName: 'output' };

  /* ══════════════════════════════════════════════════════════════
     MODE SWITCHER
     ══════════════════════════════════════════════════════════════ */
  modeUploadBtn.addEventListener('click', function () {
    modeUploadBtn.classList.add('active');
    modeEditorBtn.classList.remove('active');
    modeUpload.classList.remove('hidden');
    modeEditor.classList.add('hidden');
  });

  modeEditorBtn.addEventListener('click', function () {
    modeEditorBtn.classList.add('active');
    modeUploadBtn.classList.remove('active');
    modeEditor.classList.remove('hidden');
    modeUpload.classList.add('hidden');
  });

  /* ══════════════════════════════════════════════════════════════
     TOOLBAR  —  execCommand wrappers
     ══════════════════════════════════════════════════════════════ */

  /* Simple command buttons */
  document.querySelectorAll('.tb-btn[data-cmd]').forEach(function (btn) {
    btn.addEventListener('mousedown', function (e) {
      e.preventDefault();               /* keep editor focus */
      document.execCommand(btn.dataset.cmd, false, null);
      updateToolbarState();
    });
  });

  /* Block format dropdown */
  blockFormat.addEventListener('change', function () {
    var tag = this.value;
    editorBody.focus();
    if (tag === 'pre') {
      document.execCommand('formatBlock', false, 'pre');
    } else {
      document.execCommand('formatBlock', false, tag);
    }
    updateToolbarState();
  });

  /* Insert link */
  tbLink.addEventListener('mousedown', function (e) {
    e.preventDefault();
    var sel = window.getSelection();
    var current = '';
    if (sel && sel.rangeCount) {
      var a = sel.anchorNode && sel.anchorNode.parentElement;
      if (a && a.tagName === 'A') current = a.href;
    }
    var url = window.prompt('Enter URL (leave blank to remove link):', current);
    if (url === null) return;
    if (url.trim() === '') {
      document.execCommand('unlink', false, null);
    } else {
      if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
      document.execCommand('createLink', false, url);
    }
  });

  /* Insert <hr> */
  tbHr.addEventListener('mousedown', function (e) {
    e.preventDefault();
    document.execCommand('insertHTML', false, '<hr>');
  });

  /* Clear all */
  tbClear.addEventListener('mousedown', function (e) {
    e.preventDefault();
    if (window.confirm('Clear all editor content?')) {
      editorBody.innerHTML = '<p></p>';
      editorBody.focus();
    }
  });

  /* Reflect bold/italic/underline state when cursor moves */
  function updateToolbarState() {
    var cmds = ['bold', 'italic', 'underline', 'strikeThrough',
                'superscript', 'subscript',
                'insertUnorderedList', 'insertOrderedList',
                'justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull'];
    cmds.forEach(function (cmd) {
      var btn = document.querySelector('.tb-btn[data-cmd="' + cmd + '"]');
      if (btn) btn.classList.toggle('active', document.queryCommandState(cmd));
    });
    /* Sync block format selector */
    var block = document.queryCommandValue('formatBlock').toLowerCase().replace(/[<>]/g, '');
    var opt = blockFormat.querySelector('option[value="' + block + '"]');
    blockFormat.value = opt ? block : 'p';
  }

  editorBody.addEventListener('keyup',   updateToolbarState);
  editorBody.addEventListener('mouseup', updateToolbarState);
  document.addEventListener('selectionchange', updateToolbarState);

  /* Ensure editor always has at least one block element */
  editorBody.addEventListener('keydown', function (e) {
    if (editorBody.innerHTML.trim() === '' || editorBody.innerHTML === '<br>') {
      document.execCommand('formatBlock', false, 'p');
    }
  });

  /* ══════════════════════════════════════════════════════════════
     EDITOR HTML SANITISER
     Browsers produce messy contenteditable HTML that differs from
     Mammoth's clean output. This step normalises it BEFORE the
     shared transformHtml() rules run, so both paths produce
     identical results.

     What browsers do              → What we want
     ─────────────────────────────────────────────────────────────
     <div>…</div> (line wrap)      → <p>…</p>
     <span style="font-weight:…">  → <b> (rule4 → <strong>)
     <span style="font-style:…">   → <i> (rule5 → <strong>)
     <span style="text-decoration:underline"> → <u> (rule6 removes)
     <span style="text-decoration:line-through"> → <s>
     <div/p style="text-align:…">  → keep tag, keep align style
     <font …>                      → unwrap (strip tag, keep text)
     <span> with no useful style   → unwrap
     Nested/redundant <span>       → unwrap
     <br> inside block as only child → remove (empty line artifact)
     ══════════════════════════════════════════════════════════════ */
  function sanitizeEditorHtml(html) {
    var doc  = new DOMParser().parseFromString('<div id="__s">' + html + '</div>', 'text/html');
    var root = doc.getElementById('__s');

    /* 1. Convert <div> blocks → <p>  (browser wraps lines in div) */
    root.querySelectorAll('div').forEach(function (div) {
      /* Skip divs that contain other block elements — just unwrap */
      var hasBlock = div.querySelector('p,div,h1,h2,h3,h4,h5,h6,ul,ol,blockquote,pre,table');
      if (hasBlock) {
        unwrapElement(div);
      } else {
        changeTag(div, 'p');
      }
    });

    /* 2. Normalise <span> inline styles → semantic elements */
    /* Process deepest spans first (querySelectorAll is document order,
       so we reverse to handle nesting correctly) */
    var spans = Array.prototype.slice.call(root.querySelectorAll('span')).reverse();
    spans.forEach(function (span) {
      var style = (span.getAttribute('style') || '').replace(/\s/g, '').toLowerCase();
      if (!style) { unwrapElement(span); return; }

      /* font-weight: bold / 700+ */
      if (/font-weight\s*:\s*(bold|[7-9]\d{2})/.test(style)) {
        changeTag(span, 'b'); return;
      }
      /* font-style: italic */
      if (/font-style\s*:\s*italic/.test(style)) {
        changeTag(span, 'i'); return;
      }
      /* text-decoration: underline */
      if (/text-decoration[^;]*:\s*[^;]*underline/.test(style) &&
          !/line-through/.test(style)) {
        changeTag(span, 'u'); return;
      }
      /* text-decoration: line-through */
      if (/text-decoration[^;]*:\s*[^;]*line-through/.test(style)) {
        changeTag(span, 's'); return;
      }
      /* vertical-align: super */
      if (/vertical-align\s*:\s*super/.test(style)) {
        changeTag(span, 'sup'); return;
      }
      /* vertical-align: sub */
      if (/vertical-align\s*:\s*sub/.test(style)) {
        changeTag(span, 'sub'); return;
      }
      /* Any other span — strip the tag, keep content */
      unwrapElement(span);
    });

    /* 3. Strip <font> tags entirely — keep content */
    root.querySelectorAll('font').forEach(function (el) { unwrapElement(el); });

    /* 4. Remove lone <br> that is the only child of a block
          (artifacts from empty lines / pressing Enter) */
    root.querySelectorAll('p, li, td, th, h2, h3, h4, h5, h6').forEach(function (el) {
      if (el.childNodes.length === 1 && el.firstChild.nodeName === 'BR') {
        el.removeChild(el.firstChild);
      }
    });

    /* 5. Remove completely empty <p> tags */
    root.querySelectorAll('p').forEach(function (el) {
      if (el.textContent.trim() === '' && !el.querySelector('br,img')) {
        el.parentNode && el.parentNode.removeChild(el);
      }
    });

    /* 6. Strip ALL style attributes from every element —
          output must be pure HTML tags with no inline styles */
    root.querySelectorAll('[style]').forEach(function (el) {
      el.removeAttribute('style');
    });

    /* 7. Strip any other presentational attributes browsers add */
    root.querySelectorAll('[color],[face],[size],[bgcolor],[align],[valign],[width],[height],[cellpadding],[cellspacing],[border]').forEach(function (el) {
      ['color','face','size','bgcolor','align','valign','width','height','cellpadding','cellspacing','border'].forEach(function (attr) {
        el.removeAttribute(attr);
      });
    });

    /* 8. Strip class attributes added by browser */
    root.querySelectorAll('[class]').forEach(function (el) {
      el.removeAttribute('class');
    });

    return root.innerHTML;
  }

  /* ══════════════════════════════════════════════════════════════
     EDITOR → HTML CONVERSION
     ══════════════════════════════════════════════════════════════ */
  editorConvertBtn.addEventListener('click', function () {
    var raw = editorBody.innerHTML.trim();
    if (!raw || raw === '<p></p>' || raw === '<p><br></p>' || raw === '<br>') {
      showStatus(editorStatus, 'Editor is empty — type some content first.', 'err');
      return;
    }
    editorStatus.classList.add('hidden');

    /* Step 1: normalise browser contenteditable noise → clean semantic HTML
       Step 2: run the same transform rules as the DOCX path               */
    var sanitized   = sanitizeEditorHtml(raw);
    var transformed = transformHtml(sanitized);
    var indented    = indentHtml(transformed);

    state.htmlSource      = indented;
    state.baseName        = 'editor-output';
    fileLabel.textContent = 'Text Editor';
    previewPane.innerHTML = transformed;
    codePane.textContent  = indented;
    activateTab('preview');
    resultPanel.classList.remove('hidden');
    showStatus(editorStatus, 'Converted successfully.', 'ok');
    resultPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  });

  /* ══════════════════════════════════════════════════════════════
     HTML TRANSFORMER   (shared by DOCX + Editor paths)
     ══════════════════════════════════════════════════════════════ */
  function transformHtml(html) {
    var doc  = new DOMParser().parseFromString('<div id="__r">' + html + '</div>', 'text/html');
    var root = doc.getElementById('__r');

    rule1_headings(root);
    rule4_bToStrong(root);
    rule5_iEmToStrong(root);
    rule6_removeU(root);
    ruleD_removeStrikethrough(root);   /* NEW: remove <s>/<del>/~~text~~ entirely */
    ruleA_noStrongInHeadings(root);
    ruleB_tableScope(root);
    ruleC_mergeConsecutiveStrong(root);
    ruleE_strongWrapsAnchor(root);     /* NEW: <strong><a> not <a><strong>        */

    var s = root.innerHTML;
    s = s.replace(/<xml[\s\S]*?<\/xml>/gi, '');
    s = s.replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, '');
    s = rule3_supSubEntities(s);
    s = rule2_phoneTel(s);
    s = ruleF_fixFileUrls(s);          /* NEW: file:// href → correct web URL     */
    s = rule8_symbolEntities(s);
    s = rule10_bluecard(s);
    s = rule11_removeNbsp(s);
    return s;
  }

  /* Rule 1 — h1 → h2 */
  function rule1_headings(root) {
    root.querySelectorAll('h1').forEach(function (el) { changeTag(el, 'h2'); });
  }

  /* Rule 2 — phone/URL → tel:/href links
     Numbers near fax-related keywords are skipped.
     Checks 60 chars BEFORE and 20 chars AFTER the number. */
  function rule2_phoneTel(html) {
    var PHONE_RE = /(\+?[\d][\d\s.\-()]{6,}[\d])/g;
    var FAX_RE   = /\bfax\b/i;
    var parts = html.split(/(<a[\s\S]*?<\/a>)/gi);
    return parts.map(function (part) {
      if (/^<a[\s\S]*<\/a>$/i.test(part)) return part;
      return part.replace(/>([\s\S]*?)</g, function (m, text) {
        var r = text.replace(PHONE_RE, function (ph, _m, offset, str) {
          var before = str.slice(Math.max(0, offset - 60), offset);
          var after  = str.slice(offset + ph.length, offset + ph.length + 20);
          if (FAX_RE.test(before) || FAX_RE.test(after)) return ph;
          return '<a href="tel:' + ph.replace(/[^\d+]/g, '') + '">' + ph + '</a>';
        });
        return '>' + r + '<';
      });
    }).join('');
  }

  /* Rule 3 — sup/sub unicode entities */
  var SUP_WHOLE = { 'SM':'&#8480;','TM':'&trade;','R':'&reg;','st':'<sup>st</sup>','nd':'<sup>nd</sup>','rd':'<sup>rd</sup>','th':'<sup>th</sup>' };
  var SUP_CHAR  = { '0':'&#x2070;','1':'&sup1;','2':'&sup2;','3':'&sup3;','4':'&#8308;','5':'&#8309;','6':'&#8310;','7':'&#8311;','8':'&#8312;','9':'&#8313;','a':'&#x1D43;','b':'&#x1D47;','c':'&#x1D9C;','d':'&#x1D48;','e':'&#x1D49;','f':'&#x1DA0;','g':'&#x1D4D;','h':'&#x02B0;','i':'&#x2071;','j':'&#x02B2;','k':'&#x1D4F;','l':'&#x02E1;','m':'&#x1D50;','n':'&#x207F;','o':'&#x1D52;','p':'&#x1D56;','r':'&#x02B3;','s':'&#x02E2;','t':'&#x1D57;','u':'&#x1D58;','v':'&#x1D5B;','w':'&#x02B7;','x':'&#x02E3;','y':'&#x02B8;','z':'&#x1DBB;','+':'&#x207A;','-':'&#x207B;','=':'&#x207C;','(':'&#x207D;',')':'&#x207E;' };
  var SUB_CHAR  = { '0':'&#x2080;','1':'&#x2081;','2':'&#x2082;','3':'&#x2083;','4':'&#x2084;','5':'&#x2085;','6':'&#x2086;','7':'&#x2087;','8':'&#x2088;','9':'&#x2089;','a':'&#x2090;','e':'&#x2091;','o':'&#x2092;','x':'&#x2093;','h':'&#x2095;','k':'&#x2096;','l':'&#x2097;','m':'&#x2098;','n':'&#x2099;','p':'&#x209A;','s':'&#x209B;','t':'&#x209C;','+':'&#x208A;','-':'&#x208B;','=':'&#x208C;','(':'&#x208D;',')':'&#x208E;' };

  function rule3_supSubEntities(html) {
    html = html.replace(/<sup>([\s\S]*?)<\/sup>/gi, function (m, inner) {
      var t = inner.trim();
      if (SUP_WHOLE[t]) return SUP_WHOLE[t];
      var out = t.split('').map(function (c) { return SUP_CHAR[c] || SUP_CHAR[c.toLowerCase()] || c; }).join('');
      return out !== t ? out : m;
    });
    html = html.replace(/<sub>([\s\S]*?)<\/sub>/gi, function (m, inner) {
      var t   = inner.trim();
      var out = t.split('').map(function (c) { return SUB_CHAR[c] || SUB_CHAR[c.toLowerCase()] || c; }).join('');
      return out !== t ? out : m;
    });
    return html;
  }

  /* Rule 4 — <b> → <strong> */
  function rule4_bToStrong(root) {
    root.querySelectorAll('b').forEach(function (el) { changeTag(el, 'strong'); });
  }

  /* Rule 5 — <i>: short → <strong>, long → unwrap; <em> kept as-is (semantic emphasis) */
  function rule5_iEmToStrong(root) {
    root.querySelectorAll('i').forEach(function (el) {
      el.textContent.trim().length > 120 ? unwrapElement(el) : changeTag(el, 'strong');
    });
    /* <em> is semantically correct emphasis — leave it unchanged */
  }

  /* Rule 6 — remove <u> */
  function rule6_removeU(root) {
    root.querySelectorAll('u').forEach(function (el) { unwrapElement(el); });
  }

  /* Rule 8 — symbol entities */
  function rule8_symbolEntities(html) {
    html = html.replace(/®/g,'\u00AE').replace(/™/g,'\u2122').replace(/©/g,'\u00A9');
    html = html.replace(/®/g,'&reg;').replace(/™/g,'&trade;').replace(/©/g,'&copy;');
    html = html.replace(/\u2018/g,'&lsquo;').replace(/\u2019/g,'&rsquo;');
    html = html.replace(/\u201C/g,'&ldquo;').replace(/\u201D/g,'&rdquo;');
    html = html.replace(/(\w)'(\w)/g,'$1&apos;$2');
    html = html.replace(/\(R\)/gi,'&reg;').replace(/\(TM\)/gi,'&trade;').replace(/\(C\)/gi,'&copy;');
    return html;
  }

  /* Rule 10 — Bluecard® */
  function rule10_bluecard(html) {
    return html.split(/(<a[\s\S]*?<\/a>)/gi).map(function (p) {
      return /^<a[\s\S]*<\/a>$/i.test(p) ? p : p.replace(/\bBluecard\b(?!&reg;)/gi, function (m) { return m + '&reg;'; });
    }).join('');
  }

  /* Rule 11 — remove &nbsp; */
  function rule11_removeNbsp(html) {
    return html.replace(/&nbsp;/gi,' ').replace(/\u00A0/g,' ').replace(/ {2,}/g,' ');
  }

  /* Rule A — no <strong> inside headings */
  function ruleA_noStrongInHeadings(root) {
    root.querySelectorAll('h2 strong,h3 strong,h4 strong,h5 strong,h6 strong').forEach(function (el) { unwrapElement(el); });
  }

  /* Rule C — fix consecutive & nested <strong> tags
     1. Nested:    <strong><strong>text</strong></strong> → <strong>text</strong>
     2. Adjacent:  <strong>A</strong><strong>B</strong>   → <strong>AB</strong>
        (also handles a whitespace-only text node between them)             */
  function ruleC_mergeConsecutiveStrong(root) {
    /* Pass 1 — unwrap <strong> nested directly inside another <strong> */
    root.querySelectorAll('strong strong').forEach(function (inner) {
      unwrapElement(inner);
    });

    /* Pass 2 — merge adjacent <strong> siblings */
    root.querySelectorAll('strong').forEach(function (el) {
      var next = el.nextSibling;
      var gap  = null;
      if (next && next.nodeType === 3 && next.nodeValue.trim() === '') {
        gap  = next;
        next = next.nextSibling;
      }
      if (next && next.nodeName === 'STRONG') {
        if (gap) el.appendChild(gap);
        while (next.firstChild) el.appendChild(next.firstChild);
        next.parentNode.removeChild(next);
      }
    });
  }

  /* ─────────────────────────────────────────────────────────────
     Rule D — Remove strikethrough text entirely
     Removes the element AND its inner text content completely.
     Handles: <s>, <del>, <strike>
     Also handles markdown-style ~~text~~ in text nodes.
  ───────────────────────────────────────────────────────────── */
  function ruleD_removeStrikethrough(root) {
    /* Remove <s>, <del>, <strike> elements and all their content */
    root.querySelectorAll('s, del, strike').forEach(function (el) {
      el.parentNode.removeChild(el);
    });

    /* Remove markdown-style ~~text~~ from text nodes */
    var walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null, false);
    var textNodes = [];
    var node;
    while ((node = walker.nextNode())) textNodes.push(node);
    textNodes.forEach(function (tn) {
      if (/~~[\s\S]*?~~/.test(tn.nodeValue)) {
        tn.nodeValue = tn.nodeValue.replace(/~~[\s\S]*?~~/g, '');
      }
    });
  }

  /* ─────────────────────────────────────────────────────────────
     Rule E — <strong> must wrap <a>, never the other way around
     Incorrect: <a href="..."><strong>text</strong></a>
     Correct:   <strong><a href="...">text</a></strong>
  ───────────────────────────────────────────────────────────── */
  function ruleE_strongWrapsAnchor(root) {
    /* Case: <a> contains a <strong> as its only meaningful child
       → flip so <strong> is the outer element              */
    root.querySelectorAll('a > strong').forEach(function (strong) {
      var anchor = strong.parentNode;
      if (anchor.nodeName !== 'A') return;

      /* Check the anchor has no other non-whitespace children */
      var otherContent = Array.prototype.slice.call(anchor.childNodes).filter(function (n) {
        return n !== strong && !(n.nodeType === 3 && n.nodeValue.trim() === '');
      });
      if (otherContent.length > 0) return; /* mixed content — leave alone */

      /* Flip: replace <a><strong>text</strong></a>
               with   <strong><a>text</a></strong>  */
      var newAnchor = anchor.ownerDocument.createElement('a');
      Array.prototype.slice.call(anchor.attributes).forEach(function (attr) {
        newAnchor.setAttribute(attr.name, attr.value);
      });
      while (strong.firstChild) newAnchor.appendChild(strong.firstChild);
      strong.appendChild(newAnchor);
      anchor.parentNode.replaceChild(strong, anchor);
    });
  }

  /* ─────────────────────────────────────────────────────────────
     Rule F — Fix file:// URLs in <a href="...">
     file:// links are broken local paths. We attempt to recover
     the intended web URL from the path segments.

     Strategy (in order):
     1. Strip the file:// prefix and any leading drive/path junk.
     2. Try to parse a recognisable domain from the remaining path.
     3. If a domain-like segment is found, reconstruct as https://.
     4. Otherwise fall back to removing the href entirely (href="#")
        so the link text is kept but the broken URL is gone.
  ───────────────────────────────────────────────────────────── */
  function ruleF_fixFileUrls(html) {
    return html.replace(/(<a\s[^>]*href\s*=\s*["'])file:\/\/([^"']*)(["'][^>]*>)/gi,
      function (match, open, path, close) {
        /* Normalise backslashes and split path segments */
        var segments = path.replace(/\\/g, '/').split('/').filter(Boolean);

        /* Walk segments looking for something that looks like a hostname
           e.g. "www.google.com", "google.com", "internal.example.org"  */
        var domain = null;
        var domainIdx = -1;
        for (var i = 0; i < segments.length; i++) {
          /* Segment contains a dot and looks like a hostname */
          if (/^[a-zA-Z0-9]([a-zA-Z0-9\-]*\.)+[a-zA-Z]{2,}$/.test(segments[i])) {
            domain    = segments[i];
            domainIdx = i;
            break;
          }
        }

        var fixedUrl;
        if (domain) {
          /* Rebuild: https://domain/rest/of/path */
          var rest = segments.slice(domainIdx + 1).join('/');
          fixedUrl = 'https://' + domain + (rest ? '/' + rest : '');
        } else {
          /* No recognisable domain — remove the broken href */
          fixedUrl = '#';
        }

        return open + fixedUrl + close;
      });
  }

  /* ─────────────────────────────────────────────────────────────
     Rule B — TABLE SCOPE  (bold-driven column-heading detection)

     Runs AFTER rule4_bToStrong, so bold cells contain <strong>.
     Mammoth (and contenteditable) output every cell as <td>.

     If ALL cells in the first row are bold:
       → first row cells become <th scope="col">  (strip inner <strong>)
       → first row wrapped in <thead>
       → remaining rows wrapped in <tbody>, all cells stay <td>

     If any first-row cell is NOT bold → table left unchanged.
  ───────────────────────────────────────────────────────────── */
  function cellIsBold(cell) {
    var text = cell.textContent.trim();
    if (!text) return false;
    var strongs = Array.prototype.slice.call(cell.querySelectorAll('strong'));
    if (!strongs.length) return false;
    return strongs.map(function (s) { return s.textContent; }).join('').trim() === text;
  }

  function stripStrong(cell) {
    Array.prototype.slice.call(cell.querySelectorAll('strong')).forEach(function (s) { unwrapElement(s); });
  }

  function ruleB_tableScope(root) {
    root.querySelectorAll('table').forEach(function (table) {
      var rows = Array.prototype.slice.call(table.querySelectorAll('tr'));
      if (!rows.length) return;

      var headCells = Array.prototype.slice.call(rows[0].querySelectorAll('td, th'));
      if (!headCells.length || !headCells.every(cellIsBold)) return;

      /* Promote first row to <thead><th scope="col"> */
      headCells.forEach(function (cell) {
        stripStrong(cell);
        var th = changeTag(cell, 'th');
        th.setAttribute('scope', 'col');
      });
      var thead = table.ownerDocument.createElement('thead');
      thead.appendChild(rows[0]); /* appendChild re-parents; no need to removeChild first */

      /* Remaining rows → <tbody>, all cells stay <td> */
      var bodyRows = rows.slice(1);
      var tbody    = table.ownerDocument.createElement('tbody');
      bodyRows.forEach(function (tr) {
        Array.prototype.slice.call(tr.querySelectorAll('th')).forEach(function (th) { changeTag(th, 'td'); });
        tr.parentNode.removeChild(tr);
        tbody.appendChild(tr);
      });

      while (table.firstChild) table.removeChild(table.firstChild);
      table.appendChild(thead);
      if (bodyRows.length) table.appendChild(tbody);
    });
  }

  /* ══ DOM helpers ══════════════════════════════════════════════ */
  function changeTag(el, tag) {
    var n = el.ownerDocument.createElement(tag);
    Array.prototype.slice.call(el.attributes).forEach(function (a) { n.setAttribute(a.name, a.value); });
    while (el.firstChild) n.appendChild(el.firstChild);
    el.parentNode.replaceChild(n, el);
    return n;
  }
  function unwrapElement(el) {
    var p = el.parentNode;
    while (el.firstChild) p.insertBefore(el.firstChild, el);
    p.removeChild(el);
  }

  /* ══════════════════════════════════════════════════════════════
     HTML INDENTER
     ══════════════════════════════════════════════════════════════ */
  var INLINE_TAGS    = new Set(['a','abbr','acronym','bdo','big','br','button','cite','code','dfn','em','img','input','kbd','label','map','object','output','q','samp','select','small','span','strong','sub','sup','textarea','time','tt','var','mark','s','del','ins']);
  var SELF_CLOSE     = new Set(['area','base','br','col','embed','hr','img','input','link','meta','param','source','track','wbr']);
  var GAP_AFTER_CLOSE= new Set(['h2','h3','h4','h5','h6','p','ul','ol','li','table','div','section','article','blockquote','pre','hr']);

  function indentHtml(raw) {
    var IND = '  ', TAG_RE = /<\/?[^>]+>/g, tokens = [], last = 0, m;
    TAG_RE.lastIndex = 0;
    while ((m = TAG_RE.exec(raw)) !== null) {
      var b = raw.slice(last, m.index).replace(/\s+/g,' ').trim();
      if (b) tokens.push({ t:'text', v:b });
      tokens.push({ t:'tag', v:m[0] });
      last = m.index + m[0].length;
    }
    var tail = raw.slice(last).replace(/\s+/g,' ').trim();
    if (tail) tokens.push({ t:'text', v:tail });

    var lines = [], depth = 0;
    function pad() { return IND.repeat(Math.max(0, depth)); }
    function name(s) { var x = s.match(/<\/?([a-zA-Z][a-zA-Z0-9]*)/); return x ? x[1].toLowerCase() : ''; }
    function isSC(s, n) { return /\/>$/.test(s) || SELF_CLOSE.has(n); }

    for (var i = 0; i < tokens.length; i++) {
      var tok = tokens[i], v = tok.v, n = tok.t === 'tag' ? name(v) : '';
      if (tok.t === 'text') { lines.push(pad() + v); continue; }
      var closing = /^<\//.test(v), sc = isSC(v, n), inl = INLINE_TAGS.has(n);
      if (closing) {
        if (!inl) depth = Math.max(0, depth - 1);
        lines.push(pad() + v);
        if (!inl && GAP_AFTER_CLOSE.has(n)) lines.push('');
      } else if (sc) {
        lines.push(pad() + v);
        if (GAP_AFTER_CLOSE.has(n)) lines.push('');
      } else if (inl) {
        var run = v;
        while (i + 1 < tokens.length) {
          var nx = tokens[i + 1];
          if (nx.t === 'text') { run += nx.v; i++; }
          else if (nx.t === 'tag' && /^<\//.test(nx.v) && name(nx.v) === n) { run += nx.v; i++; break; }
          else break;
        }
        lines.push(pad() + run);
      } else {
        lines.push(pad() + v);
        depth++;
      }
    }
    return lines.join('\n');
  }

  /* ══════════════════════════════════════════════════════════════
     FILE PROCESSING
     ══════════════════════════════════════════════════════════════ */
  function processFile(file) {
    if (!file) return;
    var ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'docx')                          processDocx(file);
    else if (ext === 'xlsx'||ext==='xls'||ext==='csv') processExcel(file);
    else showStatus(uploadStatus, 'Unsupported file. Please use .docx, .xlsx, .xls, or .csv.', 'err');
  }

  /* DOCX via Mammoth */
  function processDocx(file) {
    state.baseName = file.name.replace(/\.docx$/i, '');
    showStatus(uploadStatus, 'Converting \u201C' + file.name + '\u201D\u2026', 'ok');
    resultPanel.classList.add('hidden');
    var reader = new FileReader();
    reader.onload = function (e) {
      mammoth.convertToHtml({ arrayBuffer: e.target.result })
        .then(function (result) {
          var transformed = transformHtml(result.value);
          var indented    = indentHtml(transformed);
          state.htmlSource      = indented;
          fileLabel.textContent = file.name;
          previewPane.innerHTML = transformed;
          codePane.textContent  = indented;
          activateTab('preview');
          resultPanel.classList.remove('hidden');
          showStatus(uploadStatus, 'Converted successfully.', 'ok');
          runPostConversionTests(indented, file.name);
        })
        .catch(function (err) { showStatus(uploadStatus, 'Conversion failed: ' + (err.message || err), 'err'); });
    };
    reader.onerror = function () { showStatus(uploadStatus, 'Could not read the file.', 'err'); };
    reader.readAsArrayBuffer(file);
  }

  /* Excel / CSV via SheetJS */
  function escHtml(s) {
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  function processExcel(file) {
    state.baseName = file.name.replace(/\.[^.]+$/, '');
    showStatus(uploadStatus, 'Converting \u201C' + file.name + '\u201D\u2026', 'ok');
    resultPanel.classList.add('hidden');
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb  = XLSX.read(e.target.result, { type: 'array' });
        var html = '';
        wb.SheetNames.forEach(function (sName) {
          var rows = XLSX.utils.sheet_to_json(wb.Sheets[sName], { header:1, defval:'' });
          var ne   = rows.filter(function (r) { return r.some(function (c) { return c !== '' && c != null; }); });
          if (!ne.length) return;
          var maxC = 0;
          ne.forEach(function (r) { if (r.length > maxC) maxC = r.length; });
          html += '<h2>' + escHtml(sName) + '</h2>\n<table>\n  <tbody>\n';
          ne.forEach(function (row) {
            html += '    <tr>\n';
            for (var c = 0; c < maxC; c++)
              html += '      <td>' + escHtml(row[c] !== undefined ? row[c] : '') + '</td>\n';
            html += '    </tr>\n';
          });
          html += '  </tbody>\n</table>\n\n';
        });
        html = html.trim();
        state.htmlSource      = html;
        fileLabel.textContent = file.name;
        previewPane.innerHTML = html;
        codePane.textContent  = html;
        activateTab('preview');
        resultPanel.classList.remove('hidden');
        showStatus(uploadStatus, 'Converted successfully.', 'ok');
        runPostConversionTests(html, file.name);
      } catch (err) { showStatus(uploadStatus, 'Conversion failed: ' + (err.message || err), 'err'); }
    };
    reader.onerror = function () { showStatus(uploadStatus, 'Could not read the file.', 'err'); };
    reader.readAsArrayBuffer(file);
  }

  /* ══ UI helpers ═══════════════════════════════════════════════ */
  function showStatus(el, msg, cls) {
    el.textContent = msg;
    el.className   = 'status-bar ' + cls;
    el.classList.remove('hidden');
  }

  function activateTab(name) {
    tabs.forEach(function (b) { b.classList.toggle('active', b.dataset.tab === name); });
    previewPane.classList.toggle('hidden', name !== 'preview');
    codePane.classList.toggle('hidden',   name !== 'code');
  }

  function showToast() {
    toast.classList.remove('hidden');
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { toast.classList.add('hidden'); }, 2000);
  }

  function fallbackCopy(text) {
    var ta = document.createElement('textarea');
    ta.value = text; ta.style.cssText = 'position:fixed;opacity:0';
    document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); showToast(); } catch (e) { alert('Copy failed — use the HTML code tab.'); }
    document.body.removeChild(ta);
  }

  /* ══ Events ═══════════════════════════════════════════════════ */
  fileInput.addEventListener('change', function () {
    if (this.files && this.files[0]) processFile(this.files[0]);
    this.value = '';
  });

  dropZone.addEventListener('dragenter', function (e) { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragover',  function (e) { e.preventDefault(); e.dataTransfer.dropEffect='copy'; dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', function (e) { if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over'); });
  dropZone.addEventListener('drop', function (e) {
    e.preventDefault(); dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]);
  });

  tabs.forEach(function (btn) { btn.addEventListener('click', function () { activateTab(btn.dataset.tab); }); });

  dlBtn.addEventListener('click', function () {
    if (!state.htmlSource) return;
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([state.htmlSource], { type: 'text/plain;charset=utf-8' }));
    a.download = state.baseName + '.html';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
  });

  copyBtn.addEventListener('click', function () {
    if (!state.htmlSource) return;
    navigator.clipboard && navigator.clipboard.writeText
      ? navigator.clipboard.writeText(state.htmlSource).then(showToast).catch(function () { fallbackCopy(state.htmlSource); })
      : fallbackCopy(state.htmlSource);
  });

  resetBtn.addEventListener('click', function () {
    state.htmlSource = ''; state.baseName = 'output';
    fileInput.value = ''; previewPane.innerHTML = ''; codePane.textContent = '';
    resultPanel.classList.add('hidden');
    uploadStatus.classList.add('hidden');
    editorStatus.classList.add('hidden');
    /* Hide test report button when result is cleared */
    var rptBtn = document.getElementById('test-report-btn');
    if (rptBtn) rptBtn.classList.add('hidden');
  });

  /* ══════════════════════════════════════════════════════════════
     POST-CONVERSION TEST ENGINE
     Runs a suite of checks against the converted HTML output and
     stores results so the user can download a full report.
     ══════════════════════════════════════════════════════════════ */

  /* Normalise whitespace for robust comparisons */
  function normStr(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }

  /* All test definitions — input is always the CONVERTED output HTML */
  function buildTestSuite(convertedHtml, sourceName) {
    var h = convertedHtml;

    return [
      /* ── Rule 1: no <h1> in output ─────────────────────────── */
      { id:'R1', rule:'Rule 1', name:'No <h1> tags in output (h1 → h2)',
        pass: !/<h1[\s>]/i.test(h) },

      /* ── Rule 4: no <b> tags ────────────────────────────────── */
      { id:'R4', rule:'Rule 4', name:'No <b> tags in output (b → strong)',
        pass: !/<b[\s>]/i.test(h) },

      /* ── Rule 5: no <i> or <em> tags ───────────────────────── */
      { id:'R5', rule:'Rule 5', name:'No <i> or <em> tags in output',
        pass: !/<(i|em)[\s>]/i.test(h) },

      /* ── Rule 6: no <u> tags ────────────────────────────────── */
      { id:'R6', rule:'Rule 6', name:'No <u> tags in output (underline removed)',
        pass: !/<u[\s>]/i.test(h) },

      /* ── Rule 7: no XML or MSO conditional comments ─────────── */
      { id:'R7a', rule:'Rule 7', name:'No <xml>...</xml> blocks in output',
        pass: !/<xml[\s\S]*?<\/xml>/i.test(h) },
      { id:'R7b', rule:'Rule 7', name:'No MSO conditional comments in output',
        pass: !/<!--\[if[\s\S]*?<!\[endif\]-->/i.test(h) },

      /* ── Rule 8: symbol entities converted ─────────────────── */
      { id:'R8a', rule:'Rule 8', name:'Literal © replaced with &copy; entity',
        pass: !/©/.test(h) },
      { id:'R8b', rule:'Rule 8', name:'Literal ® replaced with &reg; entity',
        pass: !/®/.test(h) },
      { id:'R8c', rule:'Rule 8', name:'Literal ™ replaced with &trade; entity',
        pass: !/™/.test(h) },

      /* ── Rule 9/B: table structure ──────────────────────────── */
      { id:'R9a', rule:'Rule B', name:'Tables with bold headers use <th scope="col">',
        pass: (function () {
          /* Only fail if there is a <table> AND its first row has <strong>
             cells but no <th scope="col"> — i.e. rule B should have fired */
          if (!/<table/i.test(h)) return true; /* no tables — N/A → pass */
          /* If any table has <thead> we know rule B ran correctly */
          if (/<thead/i.test(h)) return true;
          /* If there are tables but none with bold first rows, also fine */
          var doc2 = new DOMParser().parseFromString('<div>' + h + '</div>', 'text/html');
          var tables = doc2.querySelectorAll('table');
          var anyBoldFirstRow = false;
          tables.forEach(function (t) {
            var rows = t.querySelectorAll('tr');
            if (!rows.length) return;
            var cells = rows[0].querySelectorAll('td,th');
            if (cells.length && Array.prototype.every.call(cells, function (c) {
              return c.querySelector('strong') !== null;
            })) anyBoldFirstRow = true;
          });
          return !anyBoldFirstRow; /* bold first row exists but no thead = fail */
        })() },
      { id:'R9b', rule:'Rule B', name:'No <th> without scope attribute in table output',
        pass: (function () {
          var thTags = h.match(/<th(\s[^>]*)?\/?>/gi) || [];
          return thTags.every(function (tag) { return /\bscope\s*=/i.test(tag); });
        })() },

      /* ── Rule 11: no &nbsp; ─────────────────────────────────── */
      { id:'R11', rule:'Rule 11', name:'No &nbsp; entities in output',
        pass: !/&nbsp;/i.test(h) },

      /* ── Rule A: no <strong> inside headings ────────────────── */
      { id:'RA', rule:'Rule A', name:'No <strong> directly inside heading tags',
        pass: !/<h[2-6][^>]*>\s*<strong/i.test(h) },

      /* ── Rule C: no adjacent/nested <strong> ────────────────── */
      { id:'RC', rule:'Rule C', name:'No consecutive adjacent <strong> tags',
        pass: !/<\/strong>\s*<strong/i.test(h) },
      { id:'RC2', rule:'Rule C', name:'No nested <strong><strong> in output',
        pass: !/<strong[^>]*>\s*<strong/i.test(h) },

      /* ── Rule D: no strikethrough elements ──────────────────── */
      { id:'RD', rule:'Rule D', name:'No <s>, <del>, or <strike> tags in output',
        pass: !/<(s|del|strike)[\s>]/i.test(h) },

      /* ── Rule E: <strong> wraps <a>, not reverse ─────────────── */
      { id:'RE', rule:'Rule E', name:'No <a><strong> nesting — strong must wrap a',
        pass: !/<a\s[^>]*>\s*<strong/i.test(h) },

      /* ── Rule F: no file:// URLs ─────────────────────────────── */
      { id:'RF', rule:'Rule F', name:'No file:// URLs in href attributes',
        pass: !/href\s*=\s*["']file:\/\//i.test(h) },

      /* ── General: no inline style attributes ────────────────── */
      { id:'GEN1', rule:'General', name:'No inline style= attributes in output',
        pass: !/<[^>]+\sstyle\s*=/i.test(h) },

      /* ── General: no class attributes ───────────────────────── */
      { id:'GEN2', rule:'General', name:'No class= attributes in output',
        pass: !/<[^>]+\sclass\s*=/i.test(h) },

      /* ── General: valid HTML structure ─────────────────────── */
      { id:'GEN3', rule:'General', name:'Output is non-empty',
        pass: h.trim().length > 0 },

      /* ── Source-specific: phone links use tel: protocol ──────── */
      { id:'GEN4', rule:'Rule 2', name:'All <a href="tel:"> links use tel: protocol only',
        pass: (function () {
          var telLinks = h.match(/href\s*=\s*["']tel:[^"']*["']/gi) || [];
          return telLinks.every(function (l) { return /^href\s*=\s*["']tel:[\d+]/.test(l); });
        })() },
    ];
  }

  /* Run tests against current output and show/enable the report button */
  function runPostConversionTests(convertedHtml, sourceName) {
    state.testResults = buildTestSuite(convertedHtml, sourceName);
    state.testSource  = sourceName;
    state.testHtml    = convertedHtml;

    var pass = state.testResults.filter(function (t) { return t.pass; }).length;
    var fail = state.testResults.length - pass;

    /* Show the Test Report button */
    var rptBtn = document.getElementById('test-report-btn');
    if (rptBtn) {
      rptBtn.classList.remove('hidden');
      /* Colour the button green (all pass) or amber (some fail) */
      rptBtn.className = 'btn ' + (fail === 0 ? 'btn-report-pass' : 'btn-report-fail');
      rptBtn.querySelector('.rpt-label').textContent =
        fail === 0
          ? '✓ Tests passed (' + pass + '/' + state.testResults.length + ')'
          : '⚠ ' + fail + ' test' + (fail > 1 ? 's' : '') + ' failed — View Report';
    }
  }

  /* ── Editor: run tests after conversion (with small delay to let DOM settle) ── */
  editorConvertBtn.addEventListener('click', function () {
    setTimeout(function () {
      if (state.htmlSource) {
        runPostConversionTests(state.htmlSource, 'Text Editor');
      }
    }, 50);
  });

  /* ── Report download ─────────────────────────────────────────── */
  function downloadTestReport() {
    var tests    = state.testResults || [];
    var srcName  = state.testSource  || 'Unknown';
    var convHtml = state.testHtml    || '';
    var pass     = tests.filter(function (t) { return t.pass; }).length;
    var fail     = tests.length - pass;
    var pct      = tests.length ? Math.round(pass / tests.length * 100) : 0;
    var stamp    = new Date().toLocaleString();
    var barColor = pct === 100 ? '#2d6a4f' : pct >= 70 ? '#d97706' : '#9b2c2c';

    function eh(s) {
      return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }

    /* Summary scorecard */
    var scorecard =
      '<div style="display:flex;gap:2rem;padding:1.2rem 2rem;background:#fff;border-bottom:1px solid #e2e0d8;align-items:center;flex-wrap:wrap">'
      + '<div><div style="font-size:2rem;font-weight:800;font-family:monospace;color:#1a1a18">' + tests.length + '</div><div style="font-size:.72rem;text-transform:uppercase;letter-spacing:.4px;color:#5a5952">Total Checks</div></div>'
      + '<div><div style="font-size:2rem;font-weight:800;font-family:monospace;color:#2d6a4f">' + pass + '</div><div style="font-size:.72rem;text-transform:uppercase;letter-spacing:.4px;color:#5a5952">Passed</div></div>'
      + '<div><div style="font-size:2rem;font-weight:800;font-family:monospace;color:#9b2c2c">' + fail + '</div><div style="font-size:.72rem;text-transform:uppercase;letter-spacing:.4px;color:#5a5952">Failed</div></div>'
      + '<div style="flex:1;min-width:200px">'
      + '<div style="display:flex;align-items:center;gap:.75rem;margin-bottom:4px">'
      + '<div style="flex:1;height:10px;background:#e2e0d8;border-radius:5px;overflow:hidden"><div style="height:100%;width:' + pct + '%;background:' + barColor + ';border-radius:5px"></div></div>'
      + '<span style="font-size:.85rem;font-weight:700;color:' + barColor + ';font-family:monospace">' + pct + '%</span></div>'
      + '<div style="font-size:.72rem;color:#5a5952;text-transform:uppercase;letter-spacing:.3px">Pass rate</div>'
      + '</div>'
      + '<div style="margin-left:auto;font-size:.78rem;color:#9a9890">Source: <strong style="color:#1a1a18">' + eh(srcName) + '</strong><br>Generated: ' + stamp + '</div>'
      + '</div>';

    /* Grouped test rows */
    var groups = {};
    tests.forEach(function (t) {
      if (!groups[t.rule]) groups[t.rule] = [];
      groups[t.rule].push(t);
    });

    var tableRows = '';
    Object.keys(groups).forEach(function (rule) {
      var gPass = groups[rule].filter(function (t) { return t.pass; }).length;
      var gFail = groups[rule].length - gPass;
      tableRows += '<tr style="background:#f0efe9"><td colspan="3" style="padding:.35rem .8rem;font-weight:700;font-size:.82rem">'
        + eh(rule)
        + ' <span style="font-weight:400;color:#2d6a4f;font-size:.78em">' + gPass + ' passed</span>'
        + (gFail ? ' <span style="font-weight:400;color:#9b2c2c;font-size:.78em">/ ' + gFail + ' failed</span>' : '')
        + '</td></tr>';
      groups[rule].forEach(function (t) {
        var icon   = t.pass ? '✓' : '✗';
        var colour = t.pass ? '#2d6a4f' : '#9b2c2c';
        var rowBg  = t.pass ? '' : 'background:#fff8f8';
        tableRows += '<tr style="' + rowBg + '">'
          + '<td style="text-align:center;font-weight:700;color:' + colour + ';width:32px">' + icon + '</td>'
          + '<td style="font-family:monospace;font-size:.78rem;color:#5a5952;width:60px">' + eh(t.id) + '</td>'
          + '<td style="font-size:.85rem">' + eh(t.name) + '</td>'
          + '</tr>';
      });
    });

    /* Converted HTML snippet (first 3000 chars) */
    var snippet = eh(convHtml.length > 3000 ? convHtml.slice(0, 3000) + '\n\n… (truncated)' : convHtml);

    var report =
      '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/>'
      + '<title>Conversion Test Report — ' + eh(srcName) + '</title>'
      + '<style>'
      + '*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}'
      + 'body{font-family:system-ui,-apple-system,sans-serif;background:#f7f6f3;color:#1a1a18;font-size:14px;line-height:1.6;}'
      + 'header{background:#1a1a18;color:#f0efe9;padding:1.1rem 2rem;display:flex;align-items:center;justify-content:space-between;gap:1rem;flex-wrap:wrap;}'
      + '.brand{font-size:1.05rem;font-weight:800;font-family:monospace;letter-spacing:-.3px;}'
      + '.brand span{color:#6fcf97;}'
      + 'section{max-width:900px;margin:0 auto;padding:1.5rem 2rem 3rem;}'
      + 'h2{font-size:.95rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#5a5952;margin:1.5rem 0 .6rem;padding-bottom:.35rem;border-bottom:1px solid #e2e0d8;}'
      + 'table{width:100%;border-collapse:collapse;font-size:.85rem;}'
      + 'th{background:#1a1a18;color:#f0efe9;padding:.45rem .8rem;text-align:left;font-size:.78rem;text-transform:uppercase;letter-spacing:.3px;}'
      + 'td{padding:.42rem .8rem;border-bottom:1px solid #e2e0d8;vertical-align:top;}'
      + 'pre{background:#1a1a18;color:#a8d5a2;padding:1rem 1.25rem;border-radius:8px;font-size:.78rem;line-height:1.65;overflow-x:auto;white-space:pre-wrap;word-break:break-word;margin-top:.5rem;}'
      + 'footer{text-align:center;padding:1.5rem;font-size:.75rem;color:#9a9890;border-top:1px solid #e2e0d8;}'
      + '</style></head><body>'
      + '<header>'
      + '<div class="brand">doc<span>→</span>html <span style="font-weight:400;opacity:.5;font-size:.85rem">/ conversion test report</span></div>'
      + '<span style="font-size:.75rem;opacity:.5">' + stamp + '</span>'
      + '</header>'
      + scorecard
      + '<section>'
      + '<h2>Test Results</h2>'
      + '<table><thead><tr><th>Status</th><th>ID</th><th>Check</th></tr></thead><tbody>'
      + tableRows
      + '</tbody></table>'
      + '<h2>Converted HTML Output</h2>'
      + '<pre>' + snippet + '</pre>'
      + '</section>'
      + '<footer>doc→html Converter · Post-Conversion Test Report · ' + eh(srcName) + ' · ' + stamp + '</footer>'
      + '</body></html>';

    var blob = new Blob([report], { type: 'text/html;charset=utf-8' });
    var url  = URL.createObjectURL(blob);
    var a    = document.createElement('a');
    a.href     = url;
    a.download = 'test-report-' + new Date().toISOString().slice(0, 10) + '.html';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  /* Expose for the onclick in index.html */
  window.downloadTestReport = downloadTestReport;

})();
