/**
 * RenewEQ Land Trust Document Generator
 *
 * Template-fill architecture.
 *
 * Documents are produced by injecting values into the ORIGINAL .docx files in
 * ./templates. Only the text inside <w:t> nodes is touched; every byte of page
 * setup, styles, fonts, headers, footers and embedded media is copied through
 * untouched. Output is therefore byte-identical to the source documents except
 * for the field values — which is what county recorders require.
 *
 * Do NOT reintroduce the `docx` library to rebuild these documents. Rebuilding
 * cannot reproduce the source layout exactly and was the cause of prior drift.
 */

const express = require('express');
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static('public'));

const TEMPLATE_DIR = path.join(__dirname, 'templates');
const STATE_CLAUSES = JSON.parse(
  fs.readFileSync(path.join(__dirname, 'state-clauses.json'), 'utf8')
);

// ─── DOCUMENTS ────────────────────────────────────────────────────────────────
const DOCUMENTS = {
  appt: {
    label: 'Appointment_of_Trustee',
    template: 'Appointment of Trustee - TEMPLATE.docx',
  },
  trust: {
    label: 'Trust_Agreement',
    template: 'Trust Agreement 2.0 - TEMPLATE.docx',
  },
  deedtrust: {
    label: 'Deed_to_Trustee',
    template: 'Deed To Trustee - TEMPLATE.docx',
  },
  cert: {
    label: 'Certification_of_Living_Trust',
    template: 'Certification of Living Trust - TEMPLATE.docx',
  },
};

// ─── STATE CONFIG ─────────────────────────────────────────────────────────────
// Illinois keeps its statutory citations. Every other state uses the generic
// clause set (IL-specific verbiage stripped) and supplies its own transfer-tax
// exemption citation, if it has one.
const TRANSFER_TAX_EXEMPTION = {
  Illinois: 'Section 31-45 of the Property Tax Code',
  // Add other states here as their recorders are confirmed, e.g.:
  // Florida: 'Section 201.02(7), Florida Statutes',
};

function clauseSet(state) {
  return state === 'Illinois' ? STATE_CLAUSES.illinois : STATE_CLAUSES.generic;
}

// ─── DATE HELPERS ─────────────────────────────────────────────────────────────
const MONTHS = ['January','February','March','April','May','June','July',
                'August','September','October','November','December'];

function fmtDate(raw) {                 // 2026-03-12 -> March 12, 2026
  if (!raw) return '';
  const [y, m, d] = raw.split('-');
  return `${MONTHS[parseInt(m, 10) - 1]} ${parseInt(d, 10)}, ${y}`;
}
function ordinal(n) {
  const v = n % 100;
  return n + (['th','st','nd','rd'][(v - 20) % 10] || ['th','st','nd','rd'][v] || 'th');
}
function fmtDayOf(raw) {                // 2026-03-12 -> 12th day of March, 2026
  if (!raw) return '';
  const [y, m, d] = raw.split('-');
  return `${ordinal(parseInt(d, 10))} day of ${MONTHS[parseInt(m, 10) - 1]}, ${y}`;
}
function fmtShort(raw) {                // 2026-03-12 -> 03/12/2026
  if (!raw) return '';
  const [y, m, d] = raw.split('-');
  return `${m}/${d}/${y}`;
}
function yearOf(raw) { return raw ? raw.split('-')[0] : ''; }

// ─── XML-SAFE TOKEN SUBSTITUTION ──────────────────────────────────────────────
function xmlEscape(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

const TOKEN_RE = /\{\{([A-Z_0-9]+)\}\}/g;

/**
 * Replaces {{TOKENS}} inside <w:t> text nodes only. Runs twice so that clause
 * values which themselves contain {{STATE}} resolve fully.
 */
function fillXml(xml, values) {
  const missing = new Set();

  const pass = (text, record) =>
    text.replace(TOKEN_RE, (match, key) => {
      if (!(key in values)) {
        if (record) missing.add(key);
        return match;
      }
      return xmlEscape(values[key]);
    });

  // Only touch the inside of <w:t ...>...</w:t>
  const replaceInTextNodes = (input, record) =>
    input.replace(/(<w:t(?:\s[^>]*)?>)([^<]*)(<\/w:t>)/g,
      (m, open, body, close) => {
        if (body.indexOf('{{') === -1) return m;
        return open + pass(body, record) + close;
      });

  let out = replaceInTextNodes(xml, false);   // pass 1: clause tokens
  out = replaceInTextNodes(out, true);        // pass 2: {{STATE}} etc.
  return { xml: out, missing: [...missing] };
}

async function buildDocument(templateFile, values) {
  const buf = fs.readFileSync(path.join(TEMPLATE_DIR, templateFile));
  const zip = await JSZip.loadAsync(buf);
  const docXml = await zip.file('word/document.xml').async('string');
  const { xml, missing } = fillXml(docXml, values);
  zip.file('word/document.xml', xml);
  const out = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  return { buffer: out, missing };
}

// ─── FIELD MAPPING ────────────────────────────────────────────────────────────
function buildValues(raw) {
  const state = raw.state || 'Illinois';

  const values = {
    STATE: state,
    STATE_UPPER: state.toUpperCase(),

    TRUST_NAME: raw.trustName || '',
    TRUST_DATE: fmtDate(raw.trustDate),

    TRUSTEE_NAME: raw.trustee || '',
    TRUSTEE_ADDRESS: raw.trusteeAddr || '',
    TRUSTEE_ACCEPTANCE_DATE: fmtDayOf(raw.trustDate),

    BENEFICIARY_NAME: raw.beneficiary || '',

    SIGNER_NAME: raw.signerName || '',
    SIGNER_TITLE: raw.signerTitle || 'Manager',

    GRANTOR_NAME: raw.grantor || '',
    GRANTOR_CAPACITY: raw.grantorCapacity || '',
    GRANTOR_SIGNATURE_BLOCK: raw.grantorSignature || '',
    SELLER_NAME: raw.seller || raw.grantor || '',

    PROPERTY_ADDRESS: raw.propAddress || '',
    PROPERTY_PIN: raw.pin || '',
    PROPERTY_COUNTY: raw.county || '',
    LEGAL_DESCRIPTION: raw.legalDesc || '',
    PRIOR_INSTRUMENT_REFERENCE: raw.priorInstrument || '',

    PREPARER_NAME: raw.preparerName || '',
    PREPARER_ADDRESS: raw.preparerAddr || '7210 Manatee Ave #1278',
    PREPARER_CITY_STATE_ZIP: raw.preparerCity || 'Bradenton, FL 34209',
    TAX_BILL_MAILING_ADDRESS: raw.taxBillAddr || '',

    TRANSFER_TAX_EXEMPTION_CITATION:
      raw.exemptionCitation || TRANSFER_TAX_EXEMPTION[state] || '',
    EXEMPTION_DATE: fmtShort(raw.saleDate || raw.trustDate),

    EXECUTION_DATE: fmtDate(raw.executionDate || raw.trustDate),
    EXECUTION_YEAR: yearOf(raw.executionDate || raw.trustDate),

    NOTARY_STATE: raw.notaryState || state,
    NOTARY_COUNTY: raw.notaryCounty || raw.county || '',
    NOTARY_DATE: fmtDate(raw.notaryDate || raw.executionDate || raw.trustDate),
  };
  values.NOTARY_COUNTY_UPPER = values.NOTARY_COUNTY.toUpperCase();

  // State-specific statutory clauses
  Object.assign(values, clauseSet(state));

  // If the state has no configured exemption citation, drop the whole block.
  if (!values.TRANSFER_TAX_EXEMPTION_CITATION) {
    values.IL_EXEMPT_LINE = '';
  }

  return values;
}

// ─── API ──────────────────────────────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  try {
    const raw = req.body || {};
    const values = buildValues(raw);
    const selected = (raw.docs && raw.docs.length)
      ? raw.docs
      : Object.keys(DOCUMENTS);

    const zip = new JSZip();
    const warnings = new Set();

    for (const key of selected) {
      const doc = DOCUMENTS[key];
      if (!doc) continue;
      const { buffer, missing } = await buildDocument(doc.template, values);
      missing.forEach(m => warnings.add(m));
      zip.file(`${doc.label}.docx`, buffer);
    }

    if (warnings.size) {
      console.warn('Unresolved placeholders:', [...warnings].join(', '));
    }

    const zipBuf = await zip.generateAsync({ type: 'nodebuffer' });
    const slug = (raw.trustName || 'LandTrust').replace(/[^a-zA-Z0-9]/g, '_');
    res.set({
      'Content-Type': 'application/zip',
      'Content-Disposition': `attachment; filename="${slug}_Packet.zip"`,
      'X-Unresolved-Fields': [...warnings].join(',') || 'none',
    });
    res.send(zipBuf);
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
  const ok = templates.every(t => t.present);
  res.status(ok ? 200 : 500).json({ ok, templates });
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, '0.0.0.0', () =>
  console.log(`Land Trust Generator (template-fill) listening on ${PORT}`));
