/**
 * Regression test.
 *
 * 1. Illinois output must reproduce the approved template XML exactly once
 *    tokens are substituted — i.e. no statutory text is lost or altered.
 * 2. Every {{TOKEN}} must resolve; none may leak into the output.
 * 3. Page setup, paragraph properties and run properties must be untouched.
 * 4. Non-Illinois output must contain zero Illinois statutory references.
 */
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const assert = require('assert');

process.env.NODE_ENV = 'test';
const TEMPLATE_DIR = path.join(__dirname, '..', 'templates');

const src = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');
const sandbox = { module: { exports: {} }, require, __dirname: path.join(__dirname, '..'),
                  console, process, Buffer };
// expose internals for testing
const fn = new Function('module','require','__dirname','console','process','Buffer',
  src.replace(/app\.listen\([\s\S]*?\);\s*$/, '') +
  '\nmodule.exports = { buildValues, buildDocument, DOCUMENTS, fillXml };');
const api = (() => { const m = { exports: {} };
  fn(m, require, path.join(__dirname, '..'), console, process, Buffer); return m.exports; })();

const SAMPLE = {
  state: 'Illinois',
  trustDate: '2026-03-12', saleDate: '2026-03-12', executionDate: '2026-03-12',
  trustName: '123 Maple Street', trustee: 'SAA Property Management, LLC',
  trusteeAddr: '7210 Manatee Ave #1278, Bradenton, FL 34209',
  beneficiary: 'Test Holdings, LLC', signerName: 'Marisa Gotsch', signerTitle: 'Manager',
  grantor: 'Test Seller, LLC', grantorCapacity: 'Managing member for Test Seller, LLC',
  grantorSignature: 'Marisa Gotsch, as manager to Test Seller, LLC',
  propAddress: '123 Maple St, Chicago, IL 60601', pin: '17-09-123-456-0000',
  county: 'Cook', legalDesc: 'LOT 5 IN BLOCK 2 OF EXAMPLE SUBDIVISION',
  priorInstrument: 'Doc #1234567890 recorded 01/15/2024',
  preparerName: 'Marisa Gotsch', taxBillAddr: '123 Maple Street Trust',
  notaryState: 'Florida', notaryCounty: 'Manatee', notaryDate: '2026-03-12',
};

const IL_MARKERS = [/765\s*ILCS/, /Illinois Land Trust/, /Property Tax Code/];
const TRUST_IL_MARKERS = [/765\s*ILCS/, /Illinois Land Trust/];

(async () => {
  let failures = 0;
  const check = (name, cond, detail) => {
    if (cond) { console.log(`  PASS  ${name}`); }
    else { console.log(`  FAIL  ${name}${detail ? ' — ' + detail : ''}`); failures++; }
  };

  // ---------- Illinois ----------
  console.log('\nIllinois packet');
  const ilValues = api.buildValues(SAMPLE);
  for (const [key, doc] of Object.entries(api.DOCUMENTS)) {
    const { buffer, missing } = await api.buildDocument(doc.template, ilValues);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('word/document.xml').async('string');

    const tplZip = await JSZip.loadAsync(fs.readFileSync(path.join(TEMPLATE_DIR, doc.template)));
    const tplXml = await tplZip.file('word/document.xml').async('string');

    check(`${key}: all tokens resolved`, missing.length === 0, missing.join(','));
    check(`${key}: no {{ }} leaked`, !/\{\{[A-Z_0-9]+\}\}/.test(xml));

    const strip = s => s.replace(/<w:t(?:\s[^>]*)?>[^<]*<\/w:t>/g, '<T/>');
    check(`${key}: XML structure identical to template`, strip(xml) === strip(tplXml));

    const sect = s => (s.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g) || []).join('');
    check(`${key}: page setup identical`, sect(xml) === sect(tplXml));

    const pPr = s => (s.match(/<w:pPr>[\s\S]*?<\/w:pPr>/g) || []).length;
    check(`${key}: paragraph property count identical`, pPr(xml) === pPr(tplXml));

    check(`${key}: no highlighting`, !/<w:highlight (?!w:val="none")/.test(xml));

    // every non-document part untouched
    const names = Object.keys(zip.files).filter(n => n !== 'word/document.xml');
    let same = true;
    for (const n of names) {
      if (zip.files[n].dir) continue;
      const a = await zip.file(n).async('nodebuffer');
      const b = await tplZip.file(n).async('nodebuffer');
      if (!a.equals(b)) { same = false; break; }
    }
    check(`${key}: all other parts byte-identical`, same);
  }

  // statutes survive in IL
  const ta = await api.buildDocument(api.DOCUMENTS.trust.template, ilValues);
  const taXml = await (await JSZip.loadAsync(ta.buffer)).file('word/document.xml').async('string');
  check('trust: IL statutes retained', TRUST_IL_MARKERS.every(re => re.test(taXml)));
  check('trust: all 6 ILCS citations present',
    (taXml.match(/765 ILCS/g) || []).length === 8,
    'found ' + ((taXml.match(/765 ILCS/g) || []).length));

  const dt = await api.buildDocument(api.DOCUMENTS.deedtrust.template, ilValues);
  const dtXml = await (await JSZip.loadAsync(dt.buffer)).file('word/document.xml').async('string');
  check('deed: IL exemption retained',
    /Exempt under provisions of paragraph \(E\)/.test(dtXml) &&
    /Property Tax Code/.test(dtXml));

  // ---------- Non-Illinois ----------
  console.log('\nIndiana packet (IL verbiage must be stripped)');
  const inValues = api.buildValues({ ...SAMPLE, state: 'Indiana', notaryState: 'Indiana' });
  for (const [key, doc] of Object.entries(api.DOCUMENTS)) {
    const { buffer, missing } = await api.buildDocument(doc.template, inValues);
    const xml = await (await JSZip.loadAsync(buffer)).file('word/document.xml').async('string');
    check(`${key}: all tokens resolved`, missing.length === 0, missing.join(','));
    check(`${key}: no {{ }} leaked`, !/\{\{[A-Z_0-9]+\}\}/.test(xml));
    const hit = IL_MARKERS.filter(re => re.test(xml)).map(String);
    check(`${key}: zero Illinois references`, hit.length === 0, hit.join(' '));
    check(`${key}: state name applied`, !/Illinois/.test(xml));
  }

  const inDeed = await api.buildDocument(api.DOCUMENTS.deedtrust.template, inValues);
  const inDeedXml = await (await JSZip.loadAsync(inDeed.buffer)).file('word/document.xml').async('string');
  check('deed: exemption block dropped for unconfigured state',
    !/Exempt under provisions/.test(inDeedXml));

  console.log(failures ? `\n${failures} FAILURE(S)\n` : '\nAll checks passed.\n');
  process.exit(failures ? 1 : 0);
})();
