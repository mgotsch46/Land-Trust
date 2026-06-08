const express = require('express');
const JSZip = require('jszip');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle,
  UnderlineType, TabStopType, TabStopPosition, WidthType
} = require('docx');

const app = express();
app.use(express.json());
app.use(express.static('public'));

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function fmtDate(raw) {
  if (!raw) return '___________';
  const [y, m, d] = raw.split('-');
  const mo = ['January','February','March','April','May','June','July','August',
               'September','October','November','December'];
  return `${mo[parseInt(m)-1]} ${parseInt(d)}, ${y}`;
}

function ordDay(raw) {
  if (!raw) return '___';
  const d = parseInt(raw.split('-')[2]);
  const s = ['th','st','nd','rd'];
  const v = d % 100;
  return d + (s[(v-20)%10] || s[v] || s[0]);
}

function fmtMonthYear(raw) {
  if (!raw) return '___________';
  const [y, m] = raw.split('-');
  const mo = ['January','February','March','April','May','June','July','August',
               'September','October','November','December'];
  return `${mo[parseInt(m)-1]}, ${y}`;
}

// Standard doc properties: US Letter, 1" margins, Times New Roman 12pt
const PAGE_PROPS = {
  page: {
    size: { width: 12240, height: 15840 },
    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
  }
};

function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { after: opts.spacingAfter !== undefined ? opts.spacingAfter : 120 },
    children: [new TextRun({
      text: text || '',
      bold: opts.bold || false,
      size: opts.size || 24,
      font: 'Times New Roman',
      underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined,
    })]
  });
}

function pBlank(lines = 1) {
  return Array.from({ length: lines }, () => new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text: '', size: 24, font: 'Times New Roman' })]
  }));
}

function sigLine(label, showDate = true) {
  const children = [
    new TextRun({ text: '__________________________________ ', size: 24, font: 'Times New Roman' }),
    new TextRun({ text: showDate ? '__________________________' : '', size: 24, font: 'Times New Roman' }),
  ];
  const line1 = new Paragraph({ spacing: { after: 40 }, children });
  const line2 = new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({ text: `Signature: ${label}`, size: 24, font: 'Times New Roman' }),
      new TextRun({ text: showDate ? '                   Date' : '', size: 24, font: 'Times New Roman' }),
    ]
  });
  return [line1, line2];
}

function divider() {
  return new Paragraph({
    spacing: { after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
    children: [new TextRun({ text: '', size: 24, font: 'Times New Roman' })]
  });
}

// ─── DOCUMENT BUILDERS ────────────────────────────────────────────────────────

function buildDeed(d) {
  const children = [
    p('THIS SPACE PROVIDED FOR RECORDER\'S USE ONLY:', { size: 20 }),
    p('WHEN RECORDED RETURN TO:', { size: 20 }),
    p(d.trustName, { size: 20 }),
    ...d.returnAddr.split('\n').map(line => p(line, { size: 20 })),
    divider(),
    p(''),
    p('SPECIAL WARRANTY DEED', { bold: true, center: true, size: 28 }),
    p(''),
    p('THE GRANTOR(S),'),
    p(`${d.grantor} ${d.grantorDesc}`),
    p(`for and in consideration of: ${d.consideration} grants, bargains, sells, conveys and specially warrants to the GRANTEE(S): ${d.trustName} the following described real estate, situated at ${d.propAddress}`),
    p(''),
    p(`County of: ${d.county}`),
    p(''),
    p('Legal Description:'),
    p(d.legalDesc),
    p(''),
    p(`Commonly known as: ${d.commonAddr} (the "Property")`),
    p(''),
    p('Subject to existing taxes, assessments, liens, encumbrances, covenants, conditions, restrictions, rights of way and easements of record, the Grantor hereby covenants with the Grantee(s) the following:'),
    p(''),
    p('1. Covenant Against Encumbrances: The Grantor guarantees that there are no encumbrances upon the property other than those that have been already disclosed to the Grantee.'),
    p(''),
    p('2. Covenant of Warranty: The Grantor guarantees to the Grantee and the Grantee\'s heirs, executors and administrators that they shall have full possession of the property and not be ejected from the premises.'),
    p(''),
    p(`Grantor hereby releases and waives all rights under and by virtue of the Homestead Exemption Laws of the State of ${d.state}.`),
    p(''),
    p(`Tax Parcel Number: ${d.pin}`),
    p(''),
    p('Page 1 of 2', { center: true, size: 20 }),
    p(''),
    p('Grantor Signatures:'),
    p(''),
    ...sigLine(d.grantor, true),
    p(''),
    p(`STATE OF ___________________, COUNTY OF ______________________, ss:`),
    p(''),
    p('Notary Acknowledgment:'),
    p(''),
    p('STATE OF __________________ )'),
    p('                            ) SS                              COUNTY'),
    p('OF ________________         )'),
    p(''),
    new Paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun({ text: 'I, _________________________________ the undersigned, a Notary Public in and for said County', size: 24, font: 'Times New Roman' }),
      ]
    }),
    p('                   Printed Notary Name', { size: 20 }),
    p(`and State aforesaid, DO HEREBY CERTIFY THAT ${d.grantor} are all personally known to me to be the same people whose names are subscribed to the foregoing instrument, as having executed the same, appeared before me this day in person and acknowledged that they signed, sealed, and delivered the said instrument as their free and voluntary act for the uses and purposes therein set forth including the release and waiver of the right of homestead given under by hand.`),
    p(''),
    p('_______________________________________________'),
    p('Notary Public Signature of person taking acknowledgment'),
    p(''),
    p('_________________________________'),
    p('Title (and Rank)'),
    p(''),
    p('My commission expires _____________'),
    p(''),
    p('                                              Seal          page 2 of 2', { center: true, size: 20 }),
  ];

  return new Document({
    sections: [{ properties: PAGE_PROPS, children }]
  });
}

function buildAppt(d) {
  const month = d.trustDateRaw ? ['January','February','March','April','May','June','July','August','September','October','November','December'][parseInt(d.trustDateRaw.split('-')[1])-1] : '___';
  const year = d.trustDateRaw ? d.trustDateRaw.split('-')[0] : '____';

  const children = [
    p('APPOINTMENT OF TRUSTEE', { bold: true, center: true, size: 28 }),
    p(''),
    p(`The UNDERSIGNED Beneficiary(ies) represent that they are one hundred percent of the beneficiaries to that certain trust agreement, dated on the ${d.trustDateDay} of ${month} ${year}, and which is known as ${d.trustName} and they hereby appoint ${d.trustee}, whose address is ${d.trusteeAddr} to represent all of their beneficial interests in said trust as their Trustee for said interests jointly and severally in accepting written direction from the below named beneficiaries and their successors and serving the said trust as trustee, following their written direction in the performance of such ministerial tasks as are authorized and required by the above referenced trust agreement under which said trust has been created.`),
    p(''),
    p(`In the event of the death, disability, incapacity or refusal to act of the Trustee appointed herein, and no successor has been appointed, then the duties of the Trustee shall devolve upon the below beneficiaries and on their assigns as such time, (s)he shall transfer all trust assets by Trustee Deed to said Trustee upon duly authorized direction of the beneficiaries, and any further actions taken by the above Trustee shall be personal, and not as the authorized or as the lawful holder of the Trustee powers over the aforesaid trust. So say we all, holders of 100% of the beneficial interests.`),
    p(''),
    p(`${d.beneficiary} Beneficiary of ${d.benefitPct} undivided interest`),
    p(''),
    p(''),
    p('ACCEPTANCE/RESIGNATION OF TRUSTEE', { bold: true }),
    p(''),
    p(`TO: The above beneficiaries of the above said trust. This is to advise you that as of the ${d.trustDateDay} of ${month}, ${year}, I hereby accept the position of Trustee subject to the power of direction over the aforesaid trust together with my fiduciary to each and all of you. Please govern yourselves accordingly.`),
    p(''),
    p(''),
    new Paragraph({
      spacing: { after: 40 },
      children: [
        new TextRun({ text: '                                          ________________________________', size: 24, font: 'Times New Roman' }),
      ]
    }),
    new Paragraph({
      spacing: { after: 240 },
      children: [
        new TextRun({ text: '                                          Trustee', size: 24, font: 'Times New Roman' }),
      ]
    }),
    p(''),
    p('**THIS DOCUMENT IS NOT REQUIRED TO BE RECORDED IN ANY COUNTY RECORDER\'S OFFICE**', { bold: true, center: true }),
  ];

  return new Document({ sections: [{ properties: PAGE_PROPS, children }] });
}

function buildTrust(d) {
  const children = [
    p('LAND TRUST AGREEMENT', { bold: true, center: true, size: 28 }),
    p(''),
    p(`THIS LAND TRUST AGREEMENT ("Agreement") is entered into as of ${d.trustDate}, by and between:`),
    p(''),
    p(`TRUSTEE:     ${d.trustee}`),
    p(`             ${d.trusteeAddr}`),
    p(''),
    p(`BENEFICIARY: ${d.beneficiary}`),
    p(''),
    p(`This trust shall be known as the ${d.trustName}.`),
    p(''),
    p('ARTICLE I — TRUST PROPERTY', { bold: true }),
    p(''),
    p('The Trustee agrees to hold title to the following described real property (the "Trust Property"):'),
    p(''),
    p(`Property Address: ${d.propAddress}`),
    p(`Tax Parcel Number: ${d.pin}`),
    p(`Legal Description: ${d.legalDesc}`),
    p(''),
    p('ARTICLE II — NATURE OF TRUST / BARE LEGAL TITLE', { bold: true }),
    p(''),
    p('The Trustee shall hold only bare legal title to the Trust Property. The Trustee shall have no right, duty, or obligation to manage, control, use, sell, negotiate, or otherwise deal with the Trust Property except upon and pursuant to the written direction of the Beneficiary. The Trustee shall have no personal liability with respect to the Trust Property or any obligation arising in connection therewith. All obligations incurred in connection with the Trust Property shall be the sole responsibility of the Beneficiary.'),
    p(''),
    p('ARTICLE III — BENEFICIAL INTEREST', { bold: true }),
    p(''),
    p(`The Beneficiary holds ${d.benefitPct} of the beneficial interest in and to the Trust Property, including all rights to possession, use, income, and proceeds of the Trust Property. The beneficial interest may be assigned, transferred, or encumbered by the Beneficiary without the necessity of recording any instrument in the public records.`),
    p(''),
    p('ARTICLE IV — TRUSTEE POWERS', { bold: true }),
    p(''),
    p('Upon written direction of the Beneficiary, the Trustee is authorized and empowered to:'),
    p('(a) Purchase, acquire, hold, and take title to real property;'),
    p('(b) Sell, convey, exchange, or otherwise transfer real property;'),
    p('(c) Mortgage, pledge, or encumber real property;'),
    p('(d) Lease real property for any term;'),
    p('(e) Execute any and all documents necessary to carry out any of the foregoing.'),
    p(''),
    p('ARTICLE V — NO PERSONAL LIABILITY OF TRUSTEE', { bold: true }),
    p(''),
    p('The Trustee shall not be personally liable for any debt, claim, demand, judgment, or obligation of any nature whatsoever incurred in connection with the Trust Property or this Agreement. Any such obligation shall be payable solely from the Trust Property or from funds provided by the Beneficiary.'),
    p(''),
    p('ARTICLE VI — DURATION AND TERMINATION', { bold: true }),
    p(''),
    p('This trust shall continue until terminated by written agreement of the Trustee and Beneficiary, or upon the sale or disposition of all Trust Property, or twenty (20) years from the date hereof, whichever first occurs. Upon termination, the Trustee shall convey the Trust Property to the Beneficiary or as directed in writing by the Beneficiary.'),
    p(''),
    p('ARTICLE VII — GOVERNING LAW', { bold: true }),
    p(''),
    p(`This Agreement shall be governed by and construed in accordance with the laws of the State of ${d.state}.`),
    p(''),
    p('ARTICLE VIII — ENTIRE AGREEMENT', { bold: true }),
    p(''),
    p('This Agreement constitutes the entire agreement between the parties with respect to the subject matter hereof and supersedes all prior agreements, understandings, negotiations, and discussions, whether oral or written.'),
    p(''),
    p('IN WITNESS WHEREOF, the parties have executed this Agreement as of the date first written above.'),
    p(''),
    p('TRUSTEE:', { bold: true }),
    p(''),
    ...sigLine(d.trustee, true),
    p(''),
    p('BENEFICIARY:', { bold: true }),
    p(''),
    ...sigLine(d.beneficiary, true),
    p(''),
    p('STATE OF __________________ )'),
    p('                            ) SS'),
    p('COUNTY OF ________________  )'),
    p(''),
    p('The foregoing instrument was acknowledged before me this _______ day of ______________, ________, by _______________________________.'),
    p(''),
    p('_______________________________________________'),
    p('Notary Public'),
    p(''),
    p('My commission expires: _____________                        Seal'),
  ];

  return new Document({ sections: [{ properties: PAGE_PROPS, children }] });
}

function buildDeedToTrust(d) {
  const children = [
    p('THIS SPACE PROVIDED FOR RECORDER\'S USE ONLY:', { size: 20 }),
    p('WHEN RECORDED RETURN TO:', { size: 20 }),
    p(d.trustName, { size: 20 }),
    ...d.returnAddr.split('\n').map(line => p(line, { size: 20 })),
    divider(),
    p(''),
    p('DEED TO TRUSTEE', { bold: true, center: true, size: 28 }),
    p(''),
    p(`THIS DEED, made this ${d.trustDate}, by and between:`),
    p(''),
    p(`GRANTOR: ${d.grantor}, ${d.grantorDesc}`),
    p(''),
    p('and'),
    p(''),
    p(`GRANTEE: ${d.trustee}, not personally but solely as Trustee of the ${d.trustName}, dated ${d.trustDate}.`),
    p(''),
    p('WITNESSETH:'),
    p(''),
    p(`That the Grantor, for and in consideration of ${d.consideration}, the receipt and sufficiency of which are hereby acknowledged, does hereby grant, bargain, sell, and convey unto the Grantee, as Trustee as aforesaid and not individually, all that certain real property situated in ${d.county} County, ${d.state}, described as follows:`),
    p(''),
    p('Legal Description:'),
    p(d.legalDesc),
    p(''),
    p(`Commonly known as: ${d.commonAddr} (the "Property")`),
    p(''),
    p(`Tax Parcel Number: ${d.pin}`),
    p(''),
    p('Subject to existing taxes, assessments, liens, encumbrances, covenants, conditions, restrictions, rights of way and easements of record.'),
    p(''),
    p('IT IS EXPRESSLY UNDERSTOOD AND AGREED that the Grantee takes and holds title to said property as Trustee only, and not in any individual capacity. The Trustee shall have no power to manage, control, use, sell, negotiate, or otherwise deal with the Trust Property except upon and pursuant to the written direction of the Beneficiary of said trust. The Trustee shall have no personal liability with respect to the Trust Property.'),
    p(''),
    p('TO HAVE AND TO HOLD the said real property unto the Grantee, as Trustee as aforesaid, and to the successors and assigns of said Trustee forever.'),
    p(''),
    p('IN WITNESS WHEREOF, the Grantor has executed this Deed the day and year first above written.'),
    p(''),
    p('GRANTOR:', { bold: true }),
    p(''),
    ...sigLine(d.grantor, true),
    p(''),
    p('STATE OF ___________________, COUNTY OF ______________________, ss:'),
    p(''),
    p('Notary Acknowledgment:'),
    p(''),
    p('STATE OF __________________ )'),
    p('                            ) SS                              COUNTY'),
    p('OF ________________         )'),
    p(''),
    p('I, _________________________________ the undersigned, a Notary Public in and for said County'),
    p('                   Printed Notary Name', { size: 20 }),
    p(`and State aforesaid, DO HEREBY CERTIFY THAT ${d.grantor} are all personally known to me to be the same people whose names are subscribed to the foregoing instrument, as having executed the same, appeared before me this day in person and acknowledged that they signed, sealed, and delivered the said instrument as their free and voluntary act for the uses and purposes therein set forth.`),
    p(''),
    p('_______________________________________________'),
    p('Notary Public Signature of person taking acknowledgment'),
    p(''),
    p('_________________________________'),
    p('Title (and Rank)'),
    p(''),
    p('My commission expires _____________'),
    p(''),
    p('                                              Seal', { center: true, size: 20 }),
  ];

  return new Document({ sections: [{ properties: PAGE_PROPS, children }] });
}

function buildCert(d) {
  const children = [
    p('CERTIFICATION OF TRUST', { bold: true, center: true, size: 28 }),
    p(''),
    p(`The undersigned, ${d.trustee} ("Trustee"), hereby certifies the following with respect to the ${d.trustName}:`),
    p(''),
    p('1. TRUST EXISTENCE.', { bold: true }),
    p(`The ${d.trustName} was created on ${d.trustDate}, and is currently in full force and effect and has not been revoked, rescinded, or modified in any manner that would cause this Certification to be inaccurate.`),
    p(''),
    p('2. TRUSTEE.', { bold: true }),
    p(`The name of the Trustee is ${d.trustee}, whose address is ${d.trusteeAddr}. The Trustee has full authority to act on behalf of the Trust.`),
    p(''),
    p('3. TRUST PROPERTY.', { bold: true }),
    p(`The Trust holds title to the following described real property located in ${d.county} County, ${d.state}:`),
    p(''),
    p(`   Property Address: ${d.propAddress}`),
    p(`   Tax Parcel Number: ${d.pin}`),
    p(`   Legal Description: ${d.legalDesc}`),
    p(''),
    p('4. TRUSTEE POWERS.', { bold: true }),
    p('The Trustee has full power and authority under the Trust Agreement to:'),
    p('   (a) Acquire, hold, manage, and transfer real property;'),
    p('   (b) Sell, convey, exchange, or otherwise dispose of real property;'),
    p('   (c) Mortgage, pledge, or encumber real property;'),
    p('   (d) Lease real property;'),
    p('   (e) Execute any and all documents necessary or appropriate to accomplish any of the foregoing.'),
    p(''),
    p('5. NO REVOCATION.', { bold: true }),
    p('As of the date of this Certification, the Trust has not been revoked, rescinded, modified, or terminated.'),
    p(''),
    p('6. AVAILABILITY OF TRUST.', { bold: true }),
    p('A complete copy of the Trust Agreement is available for review upon request. Third parties, including title companies and lenders, may rely upon this Certification without requiring production of the full Trust Agreement.'),
    p(''),
    p('7. CERTIFICATION.', { bold: true }),
    p('The undersigned Trustee certifies under penalty of perjury that the foregoing representations are true and correct to the best of the Trustee\'s knowledge.'),
    p(''),
    p(`Dated: ${d.trustDate}`),
    p(''),
    p('TRUSTEE:', { bold: true }),
    p(''),
    ...sigLine(d.trustee, true),
    p(''),
    p('STATE OF __________________ )'),
    p('                            ) SS'),
    p('COUNTY OF ________________  )'),
    p(''),
    p(`The foregoing instrument was acknowledged before me this _______ day of ______________, ________, by _______________________________ as Trustee of the ${d.trustName}.`),
    p(''),
    p('_______________________________________________'),
    p('Notary Public'),
    p(''),
    p('My commission expires: _____________                        Seal'),
    p(''),
    p(`**THIS CERTIFICATION MAY BE RECORDED IN THE PUBLIC RECORDS OF ${d.county.toUpperCase()} COUNTY, ${d.state.toUpperCase()}**`, { bold: true, center: true }),
  ];

  return new Document({ sections: [{ properties: PAGE_PROPS, children }] });
}

// ─── API ENDPOINT ─────────────────────────────────────────────────────────────

app.post('/generate', async (req, res) => {
  try {
    const raw = req.body;
    const d = {
      trustDateRaw: raw.trustDate || '',
      trustDate: fmtDate(raw.trustDate),
      trustDateDay: ordDay(raw.trustDate),
      trustName: raw.trustName || '[TRUST NAME]',
      trustee: raw.trustee || '[TRUSTEE]',
      trusteeAddr: raw.trusteeAddr || '[TRUSTEE ADDRESS]',
      beneficiary: raw.beneficiary || '[BENEFICIARY]',
      benefitPct: raw.benefitPct || '100%',
      propAddress: raw.propAddress || '[PROPERTY ADDRESS]',
      county: raw.county || '[COUNTY]',
      state: raw.state || 'Florida',
      pin: raw.pin || '[TAX PARCEL NUMBER]',
      legalDesc: raw.legalDesc || '[LEGAL DESCRIPTION]',
      commonAddr: raw.commonAddr || raw.propAddress || '[PROPERTY ADDRESS]',
      grantor: raw.grantor || '[GRANTOR]',
      grantorDesc: raw.grantorDesc || 'a Single person or married couple',
      consideration: raw.consideration || 'One Dollar ($1.00) and other good and valuable consideration',
      returnAddr: raw.returnAddr || '[RETURN ADDRESS]',
    };

    const selected = raw.docs || [];
    const zip = new JSZip();

    const builders = {
      deed:      { label: 'Special_Warranty_Deed',  fn: buildDeed },
      appt:      { label: 'Appointment_of_Trustee', fn: buildAppt },
      trust:     { label: 'Trust_Agreement',         fn: buildTrust },
      deedtrust: { label: 'Deed_to_Trustee',         fn: buildDeedToTrust },
      cert:      { label: 'Certification_of_Trust',  fn: buildCert },
    };

    for (const key of selected) {
      if (builders[key]) {
        const doc = builders[key].fn(d);
        const buf = await Packer.toBuffer(doc);
        zip.file(`${builders[key].label}.docx`, buf);
      }
    }

    const zipBuf = await zip.generateAsync({ type: 'nodebuffer' });
    const trustSlug = (raw.trustName || 'LandTrust').replace(/[^a-zA-Z0-9]/g, '_');
    res.set({
      'Content-Type': 'application/zip',
      'Content-Disposition': `attachment; filename="${trustSlug}_Packet.zip"`,
    });
    res.send(zipBuf);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Land Trust Generator running on port ${PORT}`));
