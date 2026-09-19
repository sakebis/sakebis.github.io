let blocks = [], nid = 0;

const BADGE = {
  balloon: '<i class="fa-regular fa-comment"></i> balloon',
  outside: '<i class="fa-solid fa-align-right"></i> outside',
  note:    '<i class="fa-regular fa-note-sticky"></i> note'
};
const PH = {
  balloon: 'dialogue text...',
  outside: 'narration / sfx...',
  note:    'translator note...'
};

/* ── COOKIE ── only translator name + ending toggle */
function saveCookie() {
  const t = document.getElementById('translatorName').value.trim();
  const e = document.getElementById('endTog').checked;
  document.cookie = `sk_name=${encodeURIComponent(t)};max-age=31536000;path=/`;
  document.cookie = `sk_end=${e ? '1' : '0'};max-age=31536000;path=/`;
}

function loadCookie() {
  const get = k => {
    const m = document.cookie.match(new RegExp('(?:^|; )' + k + '=([^;]*)'));
    return m ? decodeURIComponent(m[1]) : null;
  };
  const name = get('sk_name'), end = get('sk_end');
  if (name !== null) {
    document.getElementById('translatorName').value = name;
    const m = document.getElementById('m_translatorName');
    if (m) m.value = name;
  }
  if (end !== null) {
    const on = end === '1';
    document.getElementById('endTog').checked = on;
    const mt = document.getElementById('m_endTog');
    if (mt) mt.checked = on;
    updateEnding();
  }
}

/* ── FILENAME ── */
function updateFilename() {
  const n = document.getElementById('manhwaName').value.trim() || '---';
  const c = String(document.getElementById('chapterNum').value || '00').padStart(2, '0');
  const t = document.getElementById('translatorName').value.trim() || '---';
  const fn = `${n} ch${c} (${t})`;
  document.getElementById('fnPreview').textContent = fn;
  const m = document.getElementById('m_fnPreview');
  if (m) m.textContent = fn;
}

/* ── BLOCKS ── */
function addBlock(type, text = '') {
  blocks.push({ id: ++nid, type, text });
  render();
  setTimeout(() => {
    const ta = document.querySelector(`[data-id="${nid}"] textarea`);
    if (ta) { ta.focus(); ta.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }
  }, 40);
}

function removeBlock(id) {
  blocks = blocks.filter(b => b.id !== id);
  render();
}

function move(id, dir) {
  const i = blocks.findIndex(b => b.id === id), j = i + dir;
  if (j < 0 || j >= blocks.length) return;
  [blocks[i], blocks[j]] = [blocks[j], blocks[i]];
  render();
}

function render() {
  const c = document.getElementById('blocks');
  c.innerHTML = '';
  document.getElementById('empty').classList.toggle('show', blocks.length === 0);
  blocks.forEach(b => {
    const row = document.createElement('div');
    row.className = 'block-row';
    row.dataset.id = b.id;
    row.innerHTML = `
      <div class="block-ctrl">
        <button class="btn-ic" onclick="move(${b.id},-1)"><i class="fa-solid fa-chevron-up"></i></button>
        <button class="btn-ic" onclick="move(${b.id},1)"><i class="fa-solid fa-chevron-down"></i></button>
        <button class="btn-ic x" onclick="removeBlock(${b.id})"><i class="fa-solid fa-xmark"></i></button>
      </div>
      <div class="block-card ${b.type}">
        <div class="block-inner">
          <div class="block-badge">${BADGE[b.type]}</div>
          <textarea placeholder="${PH[b.type]}" oninput="upd(${b.id},this)">${escHtml(b.text)}</textarea>
        </div>
      </div>`;
    c.appendChild(row);
    resize(row.querySelector('textarea'));
  });
}

function escHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function upd(id, ta) {
  const b = blocks.find(b => b.id === id);
  if (b) b.text = ta.value;
  resize(ta);
}

function resize(ta) {
  ta.style.height = 'auto';
  ta.style.height = ta.scrollHeight + 'px';
}

/* ── ENDING ── */
function updateEnding() {
  const on = document.getElementById('endTog').checked;
  document.getElementById('endBox').classList.toggle('off', !on);
  const mt = document.getElementById('m_endTog');
  if (mt) mt.checked = on;
  const mb = document.getElementById('m_endBox');
  if (mb) mb.classList.toggle('off', !on);
}

function syncEnd() {
  const on = document.getElementById('m_endTog').checked;
  document.getElementById('endTog').checked = on;
  updateEnding();
  saveCookie();
}

/* ── DRAWER ── */
function syncD(did, mid) {
  document.getElementById(did).value = document.getElementById(mid).value;
  updateFilename();
}

function openDrawer() {
  document.getElementById('m_manhwaName').value   = document.getElementById('manhwaName').value;
  document.getElementById('m_chapterNum').value   = document.getElementById('chapterNum').value;
  document.getElementById('m_translatorName').value = document.getElementById('translatorName').value;
  document.getElementById('m_endTog').checked     = document.getElementById('endTog').checked;
  updateFilename();
  document.getElementById('overlay').classList.add('open');
}

function closeDrawer() {
  document.getElementById('overlay').classList.remove('open');
}

function clearAll() {
  if (!confirm('clear all blocks?')) return;
  blocks = []; nid = 0;
  render();
  closeDrawer();
}

/* ── EXPORT DOCX ── */
async function exportDocx() {
  try {
    const n  = document.getElementById('manhwaName').value.trim()    || 'Manhwa';
    const c  = String(document.getElementById('chapterNum').value || '00').padStart(2, '0');
    const t  = document.getElementById('translatorName').value.trim() || 'Translator';
    const on = document.getElementById('endTog').checked;
    const ENDING   = 'خسته نباشی تایپیست گل ࣪ ִֶָ☾.';
    const filename = `${n} ch${c} (${t}).docx`;

    /* colors per block type */
    const C = { balloon: 'd0d0d0', outside: 'aaaaaa', note: 'cccccc', end: 'bbbbbb' };

    function xmlEsc(s) {
      return s
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    }

    /*
     * makePara — fully RTL paragraph
     *
     * Key fixes vs. the original:
     *  1. <w:rFonts w:cs="Arial"/> — "cs" targets Complex Script (Arabic/Persian)
     *     so Word actually applies the font to RTL runs.
     *  2. <w:bidi/> inside <w:pPr> tells Word the paragraph is RTL.
     *  3. <w:rtl/> inside <w:rPr> marks the run as RTL.
     *  4. <w:jc w:val="right"/> aligns text to the right.
     *  Without (1) Word ignores the font for Persian glyphs and the layout breaks.
     */
    function makePara(text, color) {
      /* handle multi-line text: split on newlines, one <w:p> per line */
      const lines = text.split('\n');
      return lines.map(line => `
<w:p>
  <w:pPr>
    <w:bidi/>
    <w:jc w:val="right"/>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:color w:val="${color}"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
      <w:rtl/>
    </w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:color w:val="${color}"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
      <w:rtl/>
    </w:rPr>
    <w:t xml:space="preserve">${xmlEsc(line)}</w:t>
  </w:r>
</w:p>
<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr></w:p>`).join('');
    }

    function emptyParas(n) {
      return Array(n).fill('<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr></w:p>').join('');
    }

    /* build body */
    let bodyXml = '';
    for (const b of blocks) {
      let txt;
      if (b.type === 'note')    txt = `*م.ت: ${b.text}`;
      else if (b.type === 'outside') txt = `*${b.text}*`;
      else                      txt = b.text;
      bodyXml += makePara(txt, C[b.type]);
    }
    if (on) {
      bodyXml += emptyParas(18);
      bodyXml += makePara(ENDING, C.end);
    }

    /* document.xml */
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${bodyXml}
    <w:sectPr>
      <w:bidi/>
      <w:textDirection w:val="btLr"/>
    </w:sectPr>
  </w:body>
</w:document>`;

    /* styles.xml — set RTL + Arial CS as document defaults */
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="24"/><w:szCs w:val="24"/>
        <w:rtl/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:bidi/>
        <w:jc w:val="right"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
</w:styles>`;

    const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:themeFontLang w:bidi="fa-IR"/>
  <w:bidi/>
</w:settings>`;

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"   Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`;

    const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    /* build zip */
    const zip = new JSZip();
    zip.file('[Content_Types].xml',          contentTypes);
    zip.file('_rels/.rels',                  rootRels);
    zip.file('word/document.xml',            documentXml);
    zip.file('word/_rels/document.xml.rels', relsXml);
    zip.file('word/styles.xml',              stylesXml);
    zip.file('word/settings.xml',            settingsXml);

    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    const url = URL.createObjectURL(blob);
    const a   = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);

  } catch (e) {
    alert('export error: ' + e.message);
    console.error(e);
  }
}

/* ── INIT ── */
window.onload = () => {
  loadCookie();
  render();
  updateFilename();
};
