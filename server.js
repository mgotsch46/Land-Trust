/**
 * RenewEQ Trust Document Generator
 *
 * Template-fill architecture. Documents are produced by substituting {{TOKEN}}
 * values inside <w:t> nodes of the original Word files and rezipping
 * entry-for-entry, so page setup, styles, fonts, headers, footers and media are
 * preserved byte-for-byte. County recorders reject documents whose margins or
 * pagination drift.
 *
 * Do NOT reintroduce a document-authoring library to rebuild these documents.
 */

const express = require('express');
const JSZip = require('jszip');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json({ limit: '5mb' }));
app.use(express.raw({ type: 'application/octet-stream', limit: '25mb' }));
app.use(express.static('public'));

const TEMPLATE_DIR = path.join(__dirname, 'templates');
const STATE_CLAUSES = JSON.parse(
  fs.readFileSync(path.join(__dirname, 'state-clauses.json'), 'utf8'));

const DOCUMENTS = {
  trust:     { label: 'Trust Agreement',                template: 'Trust Agreement 2.0 - TEMPLATE.docx' },
  appt:      { label: 'Appointment of Trustee',         template: 'Appointment of Trustee - TEMPLATE.docx' },
  deedtrust: { label: 'Deed to Trustee',                template: 'Deed To Trustee - TEMPLATE.docx' },
  cert:      { label: 'Certification of Living Trust',  template: 'Certification of Living Trust - TEMPLATE.docx' },
};
const DOC_ORDER = ['trust', 'appt', 'deedtrust', 'cert'];

// ─── STATE CONFIG ─────────────────────────────────────────────────────────────
const STATE_NAME = {
  IL: 'Illinois', MO: 'Missouri', IN: 'Indiana', FL: 'Florida',
  MI: 'Michigan', OH: 'Ohio', WI: 'Wisconsin', KS: 'Kansas', IA: 'Iowa',
};
const fullState = s => STATE_NAME[String(s || '').trim().toUpperCase()] || String(s || '').trim();

// Only states with a confirmed citation get the deed exemption stamp.
const TRANSFER_TAX_EXEMPTION = {
  Illinois: 'Section 31-45 of the Property Tax Code',
};

// States that do not recognise the land trust form use revocable trust wording.
const NO_LAND_TRUST = new Set(['Missouri']);

// Ordered so compound phrases are rewritten first — otherwise
// "Grantor Revocable Land Trust" becomes "Grantor Revocable Revocable Trust".
const RETITLE = [
  ['Grantor Revocable Land Trust', 'Grantor Revocable Trust'],
  ['Revocable Land Trust',         'Revocable Trust'],
  ['land trusts',                  'trusts'],
  ['Land Trusts',                  'Trusts'],
  ['land trust',                   'revocable trust'],
  ['Land Trust',                   'Revocable Trust'],
];
const retitle = t => RETITLE.reduce((s, [a, b]) => s.split(a).join(b), t);

// ─── DATES ────────────────────────────────────────────────────────────────────
const MONTHS = ['January','February','March','April','May','June','July',
                'August','September','October','November','December'];
function asDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const s = String(v).trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s);
  if (m) return new Date(+m[3], +m[1] - 1, +m[2]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
const ord = n => (n % 100 >= 11 && n % 100 <= 13) ? `${n}th`
  : `${n}${({ 1: 'st', 2: 'nd', 3: 'rd' })[n % 10] || 'th'}`;
const fmtLong  = d => d ? `${MONTHS[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}` : '';
const fmtDayOf = d => d ? `${ord(d.getDate())} day of ${MONTHS[d.getMonth()]}, ${d.getFullYear()}` : '';
const fmtShort = d => d ? `${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}/${d.getFullYear()}` : '';

// ─── WRITE-ON RULES ───────────────────────────────────────────────────────────
// A notary field left blank must render a visible line, not collapse to nothing.
// Lengths vary by context so alignment survives.
const BLANK_LINE = {
  NOTARY_STATE:        '_'.repeat(30),
  NOTARY_COUNTY:       '_'.repeat(30),
  NOTARY_STATE_INLINE: '_'.repeat(20),
  NOTARY_DATE:         '_'.repeat(24),
  NOTARY_STATE_UPPER:  '_'.repeat(14),
  NOTARY_COUNTY_UPPER: '_'.repeat(14),
  EXECUTION_DAY_MONTH: '_____ day of _____________',
};

const clean = v => v == null ? '' : String(v).replace(/\s*\n\s*/g, ' ').replace(/[ \t]{2,}/g, ' ').trim();
const pin   = v => (typeof v === 'number' && Number.isInteger(v)) ? String(v) : clean(v);

// ─── FIELD MAPPING ────────────────────────────────────────────────────────────
function buildValues(raw) {
  const state   = fullState(raw.state || 'Illinois');
  const rename  = NO_LAND_TRUST.has(state);
  const trustIn = clean(raw.trustName);
  const trust   = rename ? retitle(trustIn) : trustIn;

  const countyRaw = clean(raw.county);
  const isCity    = /^city of/i.test(countyRaw);
  const county    = isCity ? countyRaw : countyRaw.replace(/\s+county$/i, '');

  const trustDate = asDate(raw.trustDate);
  const saleDate  = asDate(raw.saleDate)  || trustDate;
  const execDate  = asDate(raw.executionDate) || trustDate;
  const notDate   = asDate(raw.notaryDate);

  const grantor = clean(raw.grantor);
  const signer  = clean(raw.signerName);
  const title   = clean(raw.signerTitle) || 'Managing Member';
  const ret1    = clean(raw.preparerName);
  const ret2    = clean(raw.preparerAddr);
  const priorRef = clean(raw.priorInstrument);

  const v = {
    STATE: state,
    TRUST_NAME: trust,
    TRUST_DATE: fmtLong(trustDate),
    TRUSTEE_NAME: clean(raw.trustee),
    TRUSTEE_ADDRESS: clean(raw.trusteeAddr),
    TRUSTEE_ACCEPTANCE_DATE: fmtDayOf(trustDate),
    BENEFICIARY_NAME: clean(raw.beneficiary),
    SIGNER_NAME: signer,
    SIGNER_TITLE: title,
    GRANTOR_NAME: grantor,
    GRANTOR_SIGNATURE_BLOCK: signer ? `${signer}, ${title} of ${grantor}` : '',
    ACK_SIGNER_NAME: signer || grantor,
    ACK_SIGNER_CAPACITY: signer ? `${title} of ${grantor}` : clean(raw.grantorCapacity),
    SELLER_NAME: clean(raw.seller) || grantor,
    PROPERTY_ADDRESS: clean(raw.propAddress),
    PROPERTY_PIN: pin(raw.pin),
    PROPERTY_COUNTY: county,
    LEGAL_DESCRIPTION: clean(raw.legalDesc),
    PRIOR_INSTRUMENT_LABEL: priorRef ? 'Prior instrument reference: ' : '',
    PRIOR_INSTRUMENT_REFERENCE: priorRef,
    PREPARER_NAME: ret1,
    PREPARER_ADDRESS: ret2,
    PREPARER_CITY_STATE_ZIP: clean(raw.preparerCity),
    TAX_BILL_MAILING_ADDRESS: [ret1, ret2].filter(Boolean).join(', '),
    // Trust Agreement is executed on the trust date; the deed on the sale date.
    EXECUTION_DATE: fmtLong(execDate),
    EXECUTION_YEAR: saleDate ? String(saleDate.getFullYear()) : '',
    EXECUTION_DAY_MONTH: saleDate
      ? `${ord(saleDate.getDate())} day of ${MONTHS[saleDate.getMonth()]}`
      : BLANK_LINE.EXECUTION_DAY_MONTH,
    FOOTER_DATE: fmtShort(asDate(raw.footerDate) || new Date()),
    NOTARY_STATE: clean(raw.notaryState),
    NOTARY_COUNTY: clean(raw.notaryCounty),
    NOTARY_DATE: notDate ? fmtLong(notDate) : '',
  };

  v.NOTARY_STATE_INLINE = v.NOTARY_STATE;
  v.NOTARY_STATE_UPPER  = v.NOTARY_STATE.toUpperCase();
  v.NOTARY_COUNTY_UPPER = v.NOTARY_COUNTY.toUpperCase();

  // blank notary fields become write-on lines
  for (const k of Object.keys(BLANK_LINE)) {
    if (!v[k]) v[k] = BLANK_LINE[k];
  }

  const exemption = clean(raw.exemptionCitation) || TRANSFER_TAX_EXEMPTION[state] || '';
  v.TRANSFER_TAX_EXEMPTION_CITATION = exemption;
  v.EXEMPTION_PERIOD      = exemption ? '.' : '';
  v.EXEMPTION_DATE_LABEL  = exemption ? 'Date:' : '';
  v.SELLER_LABEL          = exemption ? 'Seller: ' : '';
  v.SELLER_TRAIL1         = exemption ? ' ' : '';
  v.SELLER_TRAIL2         = exemption ? '    ' : '';
  v.EXEMPTION_DATE = exemption ? fmtShort(saleDate) : '';
  if (!exemption) {
    v.SELLER_NAME = '';                       // drop the values, not just the labels
  }

  Object.assign(v, state === 'Illinois' ? STATE_CLAUSES.illinois : STATE_CLAUSES.generic);
  if (!exemption) v.IL_EXEMPT_LINE = '';

  v._isCity = isCity;
  v._retitle = rename;
  v._dropExemption = !exemption;
  return v;
}

// ─── XML SUBSTITUTION ─────────────────────────────────────────────────────────
const xmlEscape = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const TOKEN = /\{\{([A-Z_0-9]+)\}\}/g;

function transformTextNodes(xml, fn) {
  return xml.replace(/(<w:t(?:\s[^>]*)?>)([^<]*)(<\/w:t>)/g,
    (m, open, body, close) => {
      const out = fn(body);
      return out === body ? m : open + out + close;
    });
}

function fillXml(xml, v) {
  const missing = new Set();
  const sub = (text, record) => text.replace(TOKEN, (m, k) => {
    if (!(k in v)) { if (record) missing.add(k); return m; }
    return xmlEscape(v[k]);
  });
  let out = transformTextNodes(xml, t => t.includes('{{') ? sub(t, false) : t);
  out = transformTextNodes(out, t => t.includes('{{') ? sub(t, true) : t);

  // independent cities are not counties
  if (v._isCity) {
    out = transformTextNodes(out, t =>
      t.replace('situated in the County of ', 'situated in the '));
  }
  // states without a land trust form
  if (v._retitle) {
    out = transformTextNodes(out, t =>
      (t.includes('and Trust') || t.includes('and trust')) ? retitle(t) : t);
  }
  // exemption stamp removed: clear its orphaned labels and trailing period
  if (v._dropExemption) {
    out = transformTextNodes(out, t => {
      if (t === 'Date:' || t === 'Seller: ') return '';
      return t;
    });
  }
  return { xml: out, missing: [...missing] };
}

async function buildDocument(templateFile, v) {
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(TEMPLATE_DIR, templateFile)));
  const missing = new Set();
  for (const part of ['word/document.xml', 'word/header1.xml', 'word/footer1.xml', 'word/footer3.xml']) {
    const f = zip.file(part);
    if (!f) continue;
    const r = fillXml(await f.async('string'), v);
    r.missing.forEach(m => missing.add(m));
    zip.file(part, r.xml);
  }
  return {
    buffer: await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
    missing: [...missing],
  };
}

const addrPrefix = v =>
  (v.TRUST_NAME || 'Document').replace(/\s+(Land|Revocable)\s+Trust\s*$/i, '').trim();

async function buildPacket(raw, selected) {
  const v = buildValues(raw);
  const prefix = addrPrefix(v);
  const zip = new JSZip();
  const missing = new Set();
  for (const key of (selected && selected.length ? selected : DOC_ORDER)) {
    const doc = DOCUMENTS[key];
    if (!doc) continue;
    const { buffer, missing: m } = await buildDocument(doc.template, v);
    m.forEach(x => missing.add(x));
    zip.file(`${prefix} - ${doc.label}.docx`, buffer);
  }
  return {
    buffer: await zip.generateAsync({ type: 'nodebuffer' }),
    name: (v.TRUST_NAME || 'Trust').replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_|_$/g, ''),
    missing: [...missing],
  };
}

// ─── SINGLE PACKET ────────────────────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  try {
    const raw = req.body || {};
    const { buffer, name, missing } = await buildPacket(raw, raw.docs);
    if (missing.length) console.warn('Unresolved:', missing.join(','));
    res.set({
      'Content-Type': 'application/zip',
      'Content-Disposition': `attachment; filename="${name}_Packet.zip"`,
      'X-Unresolved-Fields': missing.join(',') || 'none',
    });
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ─── BULK FROM SPREADSHEET ────────────────────────────────────────────────────
const COLS = {
  trustDate: 'Trust Date', saleDate: 'Sale Date (Deed)', trustName: 'Trust Name',
  trustee: 'Trustee', trusteeAddr: 'Trustee Address',
  preparerName: 'Return Address Line 1', preparerAddr: 'Return Address Line 2',
  beneficiary: 'Beneficiary Name', signerName: 'Director (Power of Direction)',
  propAddress: 'Property Address', county: 'County', state: 'State',
  pin: 'Tax Parcel / PIN', legalDesc: 'Legal Description',
  grantor: 'Grantor Name(s)', grantorCapacity: 'Grantor Description',
};

function rowToRaw(row) {
  const raw = {};
  for (const [k, col] of Object.entries(COLS)) raw[k] = row[col];
  return raw;
}

app.post('/generate-bulk', async (req, res) => {
  try {
    if (!req.body || !req.body.length) {
      return res.status(400).json({ error: 'No spreadsheet received.' });
    }
    const wb = XLSX.read(req.body, { type: 'buffer', cellDates: true });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
    const usable = rows.filter(r => clean(r['Trust Name']));
    if (!usable.length) {
      return res.status(400).json({ error: 'No rows with a Trust Name were found.' });
    }

    const out = new JSZip();
    const report = [];
    for (const row of usable) {
      const raw = rowToRaw(row);
      const { buffer, name, missing } = await buildPacket(raw, DOC_ORDER);
      out.file(`${name}.zip`, buffer);
      report.push({ trust: name, state: fullState(raw.state), unresolved: missing });
    }
    const buf = await out.generateAsync({ type: 'nodebuffer' });
    res.set({
      'Content-Type': 'application/zip',
      'Content-Disposition': 'attachment; filename="Trust_Packets.zip"',
      'X-Packet-Count': String(report.length),
      'X-Report': Buffer.from(JSON.stringify(report)).toString('base64'),
    });
    res.send(buf);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/health', (_req, res) => {
  const templates = Object.values(DOCUMENTS).map(d => ({
    template: d.template,
    present: fs.existsSync(path.join(TEMPLATE_DIR, d.template)),
  }));
  res.status(templates.every(t => t.present) ? 200 : 500)
     .json({ ok: templates.every(t => t.present), templates });
});

const PORT = process.env.PORT || 8080;
if (require.main === module) {
  app.listen(PORT, '0.0.0.0', () =>
    console.log(`Trust Generator (template-fill) listening on ${PORT}`));
}
module.exports = { app, buildValues, buildDocument, buildPacket, fillXml, retitle, DOCUMENTS, DOC_ORDER };
