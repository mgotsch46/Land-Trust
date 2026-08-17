/**
 * Regression suite.
 *
 * Locks in the fidelity guarantees and every correction made to these documents.
 * If a check fails, establish whether the CHECK or the OUTPUT is wrong before
 * relaxing it.
 */
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { buildValues, buildDocument, buildPacket, retitle, DOCUMENTS } = require('../server.js');

const TPL = path.join(__dirname, '..', 'templates');
let fail = 0, pass = 0;
const check = (name, cond, detail) => {
  if (cond) { pass++; console.log(`  PASS  ${name}`); }
  else { fail++; console.log(`  FAIL  ${name}${detail ? ' — ' + detail : ''}`); }
};

const BASE = {
  trustDate: '2026-08-13', saleDate: '2026-08-17',
  trustee: 'REWY Management LLC', trusteeAddr: '30 N Gould St, Ste N, Sheridan, WY 82801',
  beneficiary: 'REMO SF LLC', signerName: 'Edgardo Agrait', signerTitle: 'Managing Member',
  grantor: 'REMO SF LLC', propAddress: '339 Christian Ave, Saint Louis, MO 63147',
  pin: '5204-9-160.000', legalDesc: 'Lot 17 of Example Subdivision',
  preparerName: 'REMO SF LLC', preparerAddr: '30 N Gould St, Ste N, Sheridan, WY 82801',
};
const IL = { ...BASE, state: 'IL', trustName: '21 Dora Dr Land Trust', county: 'St. Clair County' };
const MO = { ...BASE, state: 'MO', trustName: '339 Christian Ave Land Trust', county: 'City of St. Louis' };
const MOC = { ...MO, county: 'St. Louis County' };

const textOf = async (buf, part='word/document.xml') => {
  const z = await JSZip.loadAsync(buf);
  const f = z.file(part); if (!f) return '';
  return (await f.async('string')).replace(/<[^>]+>/g, '');
};

(async () => {
  // ---------- fidelity ----------
  console.log('\nFidelity vs templates');
  for (const [key, doc] of Object.entries(DOCUMENTS)) {
    const v = buildValues(IL);
    const { buffer, missing } = await buildDocument(doc.template, v);
    const out = await JSZip.loadAsync(buffer);
    const tpl = await JSZip.loadAsync(fs.readFileSync(path.join(TPL, doc.template)));
    const ox = await out.file('word/document.xml').async('string');
    const tx = await tpl.file('word/document.xml').async('string');
    const strip = s => s.replace(/<w:t(?:\s[^>]*)?\/>/g, '<T/>')
                        .replace(/<w:t(?:\s[^>]*)?>[^<]*<\/w:t>/g, '<T/>');
    const sect = s => (s.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g) || []).join('');
    const count = (s, re) => (s.match(re) || []).length;

    check(`${key}: all tokens resolved`, missing.length === 0, missing.join(','));
    check(`${key}: no {{ }} leaked`, !/\{\{[A-Z_0-9]+\}\}/.test(ox));
    check(`${key}: XML structure identical`, strip(ox) === strip(tx));
    check(`${key}: page setup identical`, sect(ox) === sect(tx));
    check(`${key}: paragraph count identical`,
      count(ox, /<w:p[ >]/g) === count(tx, /<w:p[ >]/g));
    check(`${key}: no highlighting`, !/<w:highlight (?!w:val="none")/.test(ox));
    let same = true;
    for (const n of Object.keys(tpl.files)) {
      if (tpl.files[n].dir || n === 'word/document.xml') continue;
      if (/word\/(header|footer)\d*\.xml/.test(n)) continue;   // intentionally filled
      const a = await out.file(n).async('nodebuffer');
      const b = await tpl.file(n).async('nodebuffer');
      if (!a.equals(b)) { same = false; break; }
    }
    check(`${key}: other parts byte-identical`, same);
  }

  // ---------- Illinois ----------
  console.log('\nIllinois');
  const ilTA = await textOf((await buildDocument(DOCUMENTS.trust.template, buildValues(IL))).buffer);
  const ilDeed = await textOf((await buildDocument(DOCUMENTS.deedtrust.template, buildValues(IL))).buffer);
  check('IL: 8 ILCS citations retained', (ilTA.match(/765 ILCS/g) || []).length === 8);
  check('IL: exemption stamp present', /Exempt under provisions/.test(ilDeed) && /Property Tax Code/.test(ilDeed));
  check('IL: keeps "Land Trust" wording', /Land Trust/.test(ilTA));
  check('IL: county reads "County of St. Clair"', /County of St\. Clair,/.test(ilDeed), ilDeed.match(/situated in[^,]*/)?.[0]);

  // ---------- Missouri ----------
  console.log('\nMissouri');
  const moV = buildValues(MO);
  const moTA = await textOf((await buildDocument(DOCUMENTS.trust.template, moV)).buffer);
  const moDeed = await textOf((await buildDocument(DOCUMENTS.deedtrust.template, moV)).buffer);
  check('MO: trust renamed to Revocable Trust', moV.TRUST_NAME === '339 Christian Ave Revocable Trust', moV.TRUST_NAME);
  check('MO: zero "land trust" anywhere', !/[Ll]and\s+[Tt]rust/.test(moTA + moDeed));
  check('MO: no doubled "Revocable Revocable"', !/Revocable\s+Revocable/.test(moTA + moDeed));
  check('MO: zero ILCS citations', !/ILCS/.test(moTA));
  check('MO: no "Illinois" anywhere', !/Illinois/.test(moTA + moDeed));
  check('MO: exemption stamp dropped', !/Exempt under provisions/.test(moDeed));
  check('MO: no orphaned Date:/Seller: labels', !/Date:\s*Seller:/.test(moDeed) && !/Seller:\s{0,3}DEED/.test(moDeed));
  check('MO: independent city reads "in the City of St. Louis"',
    /situated in the City of St\. Louis,/.test(moDeed), moDeed.match(/situated in[^,]*/)?.[0]);
  const moCounty = await textOf((await buildDocument(DOCUMENTS.deedtrust.template, buildValues(MOC))).buffer);
  check('MO: normal county reads "County of St. Louis"', /County of St\. Louis,/.test(moCounty));
  check('retitle() is order-safe',
    retitle('Grantor Revocable Land Trust') === 'Grantor Revocable Trust');

  // ---------- corrections ----------
  console.log('\nDocument corrections');
  const ilAppt = await textOf((await buildDocument(DOCUMENTS.appt.template, buildValues(IL))).buffer);
  const ilCert = await textOf((await buildDocument(DOCUMENTS.cert.template, buildValues(IL))).buffer);
  check('no duplicated "Trust Trust"', !/Trust\s+Trust/.test(ilAppt + ilDeed + ilTA + ilCert));
  check('TA 1.1 name filled', !/NAME OF TRUST/.test(ilTA) && /21 Dora Dr Land Trust/.test(ilTA));
  check('TA 1.1 effective date filled', !/00\/00\/2026/.test(ilTA));
  check('TA signature block updated', !/Darin Knox/.test(ilTA) && /Edgardo Agrait, Managing Member/.test(ilTA));
  check('deed PROPERTY ADDRESS filled', /PROPERTY ADDRESS: 339 Christian Ave/.test(moDeed));
  check('deed PIN filled', /PIN#: 5204-9-160\.000/.test(moDeed));
  check('deed jurat year not 2025', !/___________, 2025/.test(ilDeed) && /___________, 2026/.test(ilDeed));
  check('deed execution date filled', /this 17th day of August, 2026/.test(ilDeed));
  check('acknowledgment names the individual',
    /Edgardo Agrait, as Managing Member of REMO SF LLC/.test(ilDeed));
  check('prior-instrument line hidden when empty', !/Prior instrument reference/.test(ilDeed));
  const withRef = await textOf((await buildDocument(DOCUMENTS.deedtrust.template,
    buildValues({ ...IL, priorInstrument: 'Doc #123' }))).buffer);
  check('prior-instrument line shown when supplied', /Prior instrument reference: Doc #123/.test(withRef));

  // ---------- notary write-on lines ----------
  console.log('\nNotary write-on lines');
  check('cert: State of has a line', /State of _{30}/.test(ilCert));
  check('cert: County of has a line', /County of _{30}/.test(ilCert));
  check('cert: date line + comma', /On _{24}, before me/.test(ilCert));
  check('cert: mid-sentence line is shorter', /State of _{20} that the foregoing/.test(ilCert));
  check('TA: State of / County of lines', /State of: _{30}/.test(ilTA) && /County of: _{30}/.test(ilTA));
  check('TA: date line', /On _{24}, before me/.test(ilTA));
  check('deed: venue lines', /STATE OF _{14}/.test(ilDeed) && /COUNTY OF _{14}/.test(ilDeed));
  const filled = await textOf((await buildDocument(DOCUMENTS.cert.template,
    buildValues({ ...IL, notaryState: 'Florida', notaryCounty: 'Duval', notaryDate: '2026-08-17' }))).buffer);
  check('supplied notary values replace the lines',
    /State of Florida/.test(filled) && /County of Duval/.test(filled) &&
    /On August 17, 2026, before me/.test(filled) &&
    !/State of _/.test(filled) && !/County of _/.test(filled) && !/On _/.test(filled));

  // ---------- header / footer ----------
  console.log('\nHeader and footer');
  const taBuf = (await buildDocument(DOCUMENTS.trust.template, buildValues({ ...IL, footerDate: '2026-08-16' }))).buffer;
  for (const part of ['word/header1.xml', 'word/footer1.xml', 'word/footer3.xml']) {
    const t = await textOf(taBuf, part);
    check(`${part}: trust name present`, /21 Dora Dr Land Trust/.test(t), t.slice(0, 60));
    check(`${part}: no placeholder`, !/ame [Oo]f Trust/.test(t));
    const z = await JSZip.loadAsync(taBuf);
    const raw = await z.file(part).async('string');
    check(`${part}: no highlighting`, !/<w:highlight (?!w:val="none")/.test(raw));
  }
  check('footer carries the date', /08\/16\/2026/.test(await textOf(taBuf, 'word/footer1.xml')));

  // ---------- packet ----------
  console.log('\nPacket assembly');
  const pk = await buildPacket(MO, null);
  const names = Object.keys((await JSZip.loadAsync(pk.buffer)).files);
  check('4 documents in the packet', names.length === 4, names.join(', '));
  check('documents prefixed by address',
    names.every(n => n.startsWith('339 Christian Ave - ')), names.join(', '));
  check('zip named for the renamed trust',
    pk.name === '339_Christian_Ave_Revocable_Trust', pk.name);

  console.log(`\n${pass} passed, ${fail} failed\n`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
