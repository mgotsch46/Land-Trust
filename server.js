const express = require('express');
const JSZip = require('jszip');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle,
  UnderlineType, TabStopType, WidthType, HeadingLevel, IndentationProps
} = require('docx');

const app = express();
app.use(express.json());
app.use(express.static('public'));

// ─── PAGE SETUP ───────────────────────────────────────────────────────────────
const PAGE_PROPS = {
  page: {
    size: { width: 12240, height: 15840 },
    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
  }
};

// ─── PARAGRAPH HELPERS ────────────────────────────────────────────────────────
function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { after: opts.after !== undefined ? opts.after : 160, line: opts.line || 276 },
    indent: opts.indent ? { left: opts.indent } : undefined,
    children: [new TextRun({
      text: text || '',
      bold: opts.bold || false,
      italics: opts.italics || false,
      size: opts.size || 24,
      font: 'Times New Roman',
    })]
  });
}

function pb() {
  return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: '', size: 24, font: 'Times New Roman' })] });
}

function pBold(text, opts = {}) { return p(text, { ...opts, bold: true }); }
function pCenter(text, opts = {}) { return p(text, { ...opts, center: true }); }
function pCenterBold(text, opts = {}) { return p(text, { ...opts, center: true, bold: true }); }

function sigPair(label) {
  return [
    new Paragraph({
      spacing: { after: 60 },
      children: [
        new TextRun({ text: '_________________________________________ ', size: 24, font: 'Times New Roman' }),
        new TextRun({ text: '  _________________________', size: 24, font: 'Times New Roman' }),
      ]
    }),
    new Paragraph({
      spacing: { after: 200 },
      children: [
        new TextRun({ text: label, size: 24, font: 'Times New Roman' }),
        new TextRun({ text: '        Date', size: 24, font: 'Times New Roman' }),
      ]
    }),
  ];
}

function sigSingle(label) {
  return [
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: '_________________________________________', size: 24, font: 'Times New Roman' })]
    }),
    new Paragraph({
      spacing: { after: 200 },
      children: [new TextRun({ text: label, size: 24, font: 'Times New Roman' })]
    }),
  ];
}

// ─── DATE HELPERS ─────────────────────────────────────────────────────────────
function fmtDate(raw) {
  if (!raw) return '___________';
  const [y, m, d] = raw.split('-');
  const mo = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return `${mo[parseInt(m)-1]} ${parseInt(d)}, ${y}`;
}
function fmtDateShort(raw) {
  if (!raw) return '__/__/____';
  const [y, m, d] = raw.split('-');
  return `${m}/${d}/${y}`;
}
function ordDay(raw) {
  if (!raw) return '___';
  const d = parseInt(raw.split('-')[2]);
  const s = ['th','st','nd','rd'];
  const v = d % 100;
  return d + (s[(v-20)%10] || s[v] || s[0]);
}
function monthName(raw) {
  if (!raw) return '___';
  const mo = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return mo[parseInt(raw.split('-')[1])-1];
}
function year(raw) { return raw ? raw.split('-')[0] : '____'; }

// ─── DOCUMENT BUILDERS ────────────────────────────────────────────────────────

function buildDeedToTrustee(d) {
  const children = [
    p('Prepared by:'),
    p('7210 Manatee Ave #1278'),
    p('Bradenton, FL 34209'),
    pb(),
    p('Mail Subsequent Tax Bills To:'),
    p(d.trustName),
    p(d.returnAddr1),
    p(d.returnAddr2),
    pb(),
    p(`Parcel ID #: ${d.pin}`),
    pb(),
    p('PRIOR DEED REFERENCE:'),
    pb(),
    pCenterBold('DEED TO TRUSTEE', { size: 26 }),
    pb(),
    p(`THIS INDENTURE WITNESSETH, that the Grantor, ${d.grantor}, ${d.grantorDesc}, for and in consideration of the sum of One Dollar ($1.00) and other good and valuable considerations in hand paid, the receipt and sufficiency of same being hereby acknowledged, does hereby grant, sell and convey unto ${d.trustee}, as Trustee under a trust agreement known as the ${d.trustName} dated ${fmtDateShort(d.trustDate)}, and any amendments thereto, the following real estate, situated in the County of ${d.county}, and State of ${d.state}, to-wit:`),
    pb(),
    p(d.legalDesc),
    pb(),
    p(`PROPERTY ADDRESS: ${d.propAddress}`),
    p('TO HAVE AND TO HOLD the said premises with the appurtenances upon the trusts and for the uses and purposes herein and in said Trust Agreement set forth.'),
    pb(),
    p('Full power and authority is hereby granted to said Trustee to improve, manage, protect and subdivide said premises or any part thereof, to dedicate parks, streets, highways or alleys and to vacate any subdivision or part thereof, and to re-subdivide said property as often as desired, to contract to sell, to grant options to purchase, to sell on any terms, to convey, either with or without consideration, to convey said premises or any part to a successor or successors in trust, and to grant to each successor or successors in trust all of the title, estate, powers, and authorities vested in said trustee; to donate, to dedicate, to mortgage, pledge or otherwise encumber said property, or any part thereof, to lease said property, or any part thereof, from time to time, in possession or reversion, by leases to commence in praesenti or in futuro, and upon any terms and for any period or periods of time, not exceeding in the case of any single demise the term of 198 years, and to renew or extend leases upon any terms and for any period or periods of time and to amend, change or modify leases and the terms and provisions thereof at any time or times hereafter, to contract to make leases and to grant options to lease and options to renew leases and options to purchase the whole or any part of the reversion and to contract respecting the manner of fixing the amount of present or future rentals, to partition or exchange said property, or any part thereof, for other real or personal property, to grant easements or charges of any kind, to release, convey or assign any right, title or interest in or about or easement appurtenant to said premises or any part thereof, and to deal with said property and every part thereof in all other ways and for such other considerations as it would be lawful for any person owning the same to deal with the same, whether similar to or different from the ways above specified, at any time or times hereafter.'),
    pb(),
    p('In no case shall any party dealing with said Trustee in relation to said premises, or to whom said premises or any part thereof shall be conveyed, contracted to be sold, leased or mortgaged by said Trustee, be obliged to see to the application of any purchase money, rent, or money borrowed or advanced on said premises, or be obliged to see to the application of any purchase money, rent, or money borrowed or advanced on said premises, or be obliged to see that the terms of this trust have been complied with, or be obliged to inquire into the necessity or expediency of any act of said trustee, or be obliged or privileged to inquire into any of the terms of said trust agreement; and every deed, trust deed, mortgage, lease or other instrument executed by said trustee in relation to said real estate shall be conclusive evidence in favor of every person relying upon or claiming under any such conveyance, lease or other instrument, (a) that at the time of the delivery thereof the trust created by this Indenture and by said trust agreement was in full force and effect, (b) that such conveyance or other instrument was executed in accordance with the trusts, conditions and limitations contained in this Indenture and in said Trust Agreement or in some amendment thereof and binding upon all beneficiaries thereunder, (c) that the Trustee was duly authorized and empowered to execute and deliver every such deed, trust deed, lease, mortgage or other instrument and (d) if the conveyance is made to a successor or successors in trust, that such successor or successors in trust have been properly appointed and are fully vested with all the title, estate, rights, powers, authorities, duties and obligations of its, his or their predecessors in trust.'),
    pb(),
    p('The interest of each and every beneficiary hereunder and of all persons claiming under them or any of them shall be only in the earnings, avails and proceeds arising from the sale or other disposition of said real estate, and such interest is hereby declared to be personal property, and no beneficiary hereunder shall have any title or interest, legal or equitable, in or to said real estate as such, but only an interest in the earnings, avails and proceeds thereof as aforesaid.'),
    pb(),
    p('If the title to any of the above lands is now or hereafter registered, the Registrar of Titles is hereby directed not to register or note in the certificate of title or duplicate thereof, or memorial, the words, "in trust", or "upon condition", or "with limitations", or words of similar import, in accordance with the statute in such case made and provided.'),
    pb(),
    p(`And the said Grantor(s) hereby expressly waive and release any and all right or benefit under and by virtue of any and all statutes of the State of ${d.state}, providing for the exemption of homesteads from sale on execution or otherwise.`),
    pb(),
    pBold('Exempt under provisions of paragraph E'),
    pBold('Section 4, Real Estate Transfer Tax Act.'),
    pb(),
    pBold(`Date: ${d.saleDate || fmtDateShort(d.trustDate)}`),
    pBold(`Seller: ${d.grantor}`),
    pb(),
    p(`IN WITNESS WHEREOF, said Grantor aforesaid has executed this instrument this ${ordDay(d.trustDate)} day of ${monthName(d.trustDate)}, ${year(d.trustDate)}.`),
    pb(),
    p(`STATE OF _____________`),
    p(`SS.                                         COUNTY OF _________________`),
    pb(),
    p(`I, the undersigned, a Notary Public, in and for said County, in the State aforesaid, do hereby certify that, ${d.grantor}, appeared before me this day in person and acknowledged that they signed, sealed and delivered the said instrument as their free and voluntary act for the uses and purposes therein set forth.`),
    pb(),
    p('Given under my hand and notarial seal this ______ day of ___________, ' + year(d.trustDate) + '.'),
    pb(),
    p('__________________________________________'),
    p('Notary Public'),
    pb(),
    pCenter('1   Of   3', { size: 20 }),
  ];

  return new Document({ sections: [{ properties: PAGE_PROPS, children }] });
}

function buildTrustAgreement(d) {
  // Build the full trust agreement with exact verbatim language, swapping only the highlighted fields
  const children = [
    pCenterBold('TRUST AGREEMENT', { size: 26 }),
    pb(),
    pCenterBold(`${d.trustName}`, { size: 26 }),
    pb(),
    p(`\tThis Trust Agreement is made by ${d.trustee}, as Trustee (the "Trustee").  The real property (the "trust estate") listed in the Schedule of Trust Property attached hereto and made part hereof shall be held in trust and administered and/or distributed as provided in this instrument. The name of the trust created by this instrument shall be the ${d.trustName}.  This agreement shall be effective immediately on acceptance by the Trustee as indicated on the attached schedule of property attached hereto.`),
    pb(),
    pCenterBold('ARTICLE ONE'),
    pCenterBold('CREATION OF TRUST'),
    pb(),
    p('\t\t1.1\tPurpose of Trust.  This trust shall be REVOCABLE.', { indent: 720 }),
    pb(),
    p('\t\tThe Grantor shall have no right or power, whether alone or in conjunction with others, in whatever capacity, to alter, amend, revoke or terminate this trust or any of the terms of this Agreement, in whole or in part, or to designate the persons who shall possess or enjoy the trust property, or the income therefrom. By this Agreement, the Grantor intends to and does hereby relinquish absolutely and forever every interest of any nature, present or future, in the trust property, as well as all possession and enjoyment of, or right to the income from, the trust property, whether directly, indirectly or constructively.'),
    pb(),
    p('\t\t The Trustee shall hold title to real property and protect, conserve, administer, and distribute the real property described in the Schedule of Trust Property (which is attached hereto and made a part of this trust instrument), and any other real property that may be hereafter subject to this trust, and the income and proceeds attributable to all such property, in accordance with the provisions of this instrument.'),
    pb(),
    p('\t\t\t1.2\tAdditions.  The Trustee, in its discretion, may accept additions of real property to this trust from any source by confirming such addition in writing to the beneficiaries.  Additions, if any, shall become part of the trust and held in accordance with the terms of this trust.'),
    pb(),
    p('\t\t\t1.3\tDeeds to Trustee.  The titling of any deed to real property in the name of the Trustee, as Trustee of this trust, or any successor Trustee of this trust, shall be deemed to be a transfer to this trust.  The Trustee and the Beneficiaries agree that when the Trustee has taken title to the property transferred to it and accepted by it, that the Trustee will hold the property subject to the trust purposes and conditions and the purposes and conditions contained in the instrument conveying the property.'),
    pb(),
    p('\t\t\t- Type of Trust.  Neither this Trust Agreement nor any actions of the Trustee or Beneficiaries shall be deemed to be, or to create, or be evidence of the existence of a corporation, de facto or de jure, or a Trust, or any other type of business trust, association, general or limited partnership, limited liability company or joint venture between or among the creator or holder of the beneficial interests hereunder or between the Trustee and said Beneficiary or between or among Beneficiaries of this Trust.  This Trust is intended to be a title holding or land trust.'),
    pb(),
    pCenterBold('ARTICLE TWO'),
    pCenterBold('BENEFICIARIES'),
    pb(),
    p(`\t\t\t2.1\tUse by Beneficiaries.  The property of this trust shall be held for the ultimate use and benefit of the beneficiary, or beneficiaries, listed in this Article Two.  The beneficiaries shall have the proportionate interests stated in paragraph 2.2 if there are multiple beneficiaries.`),
    pb(),
    p('\t\t\t2.2\tEvidence of Beneficial Interest; Certificates.  The following person(s), entity(ies), partnership(s) or Personal Property Trust(s) is/are the beneficiaries of this trust:'),
    pb(),
    p('Name and Address:'),
    pb(),
    p(`\t1: ${d.beneficiary} ${d.beneficiaryAddr}`),
    pb(),
    p(`Interest percentage: ${d.benefitPct}`),
    pb(),
    p('\t\t\t\t Each beneficiary shall be given a Certificate of Trust Beneficial Interest upon request to the Trustee indicating his/her/its ownership of a personal property interest in the trust.'),
    pb(),
    p('\t\t\t\t2.3\tOwnership.  The interest of a beneficiary of this trust shall be solely one of personal property and may be assigned as such in accordance with paragraph 2.4(d).  Legal and equitable title to the real property held by this trust shall be vested in the Trustee.'),
    pb(),
    p('\t\t\t\t2.4\tRights and Powers of Beneficiaries.  The beneficiary(ies) shall have the following rights and powers:'),
    pb(),
    p('\t\t2.4(a)\tPower of Direction.'),
    pb(),
    p('(i)\t Trustee as Owner.; Beneficiary(ies) Hold Power of Direction.  The Trustee is the sole owner of the property transferred to this trust, and, so far as third parties are concerned, has full power to deal with said property. However, the present and future beneficiary(ies) of this trust shall be the sole holder(s) of the power of direction over the title to trust property and shall have the sole right to direct the Trustee to convey or otherwise deal with the real property held by this trust.'),
    pb(),
    p('(ii)\tMultiple Beneficiaries; Unanimous Vote.  If there is more than one beneficiary, and no third-party Director is named pursuant to Article Three hereof, the beneficiary(ies) shall exercise the power of direction by unanimous vote.'),
    pb(),
    p('(iii)\tMultiple Beneficiaries; One appointed to Direct.  If there is more than one beneficiary and the beneficiaries select one of the beneficiaries to hold the power of direction, they shall select that person by majority vote and shall appoint such beneficiary in writing delivered to the Trustee.'),
    pb(),
    p('(iv)\tAppointment of Third-Party Director.  Notwithstanding the foregoing, the beneficiary(ies) shall have right to transfer the power of direction over the Trustee and to appoint another person or a Board of Directors to hold such power of direction in accordance with Paragraph 3.1.  When a third party named by the beneficiary(ies) pursuant to Article III holds such power of direction, he or she shall be called the Director.'),
    pb(),
    p('(v)\tAuthorized Written Direction.  It is understood and agreed by all present and future beneficiaries hereunder that the Trustee will deal with trust property only when authorized so to do in writing by the person(s) properly holding the power of direction, whether that written direction is by the beneficiary(ies) as provided in (i) above, by one of the beneficiaries specifically appointed to exercise the power of direction for multiple beneficiaries as provided in (ii) above, or by the Director if one is appointed by the beneficiary(ies) as provided in (iii) above, whichever holds the right to so direct at that time.'),
    pb(),
    p('\t\t\t\t\t\t\t\t\t\t\t\t\t(vi)\tTrustee Subject to Power of Direction.  The exercise of the Trustee powers in Article Four, paragraph 4.2, shall be subject to this power of direction, whether held by one or more of the beneficiaries or their appointed Director.'),
    pb(),
    p('\t\t\t\t\t\t\t\t\t\t\t\t\t(vii) If the Beneficiary is served notice of a lawsuit against him/her/it, the Beneficiary must immediately transfer his/her/its beneficial interest to the Successor Beneficiary and notify the Trustee of such transfer.'),
    pb(),
    p('\t\t\t\t2.4(b)\tRight to Proceeds and Avails.   The beneficiary or beneficiaries shall have the right to receive the net proceeds from rental or other income, mortgages, the net gain from sales, or other dispositions of the real property held by this trust in the same proportion as the shares listed for each of them in paragraph 2.2.  Any benefits to a beneficiary shall be distributed to a beneficiary at the direction of the beneficiary, and only upon the beneficiary\'s written request, which is a privilege that may be exercised only voluntarily and shall not include any involuntary exercise.'),
    pb(),
    p('\t\t\t\t2.4 (c)\tRight of Management and Control. The beneficiary or beneficiaries shall have the full power of management and control of the real property held by this trust, and of the selling, renting, and handling thereof, including the collection of rent and proceeds of sale, and hiring of property managers. The beneficiary or beneficiaries shall also control the payment of taxes, assessments, insurance, and other expenses in connection with the real property.  The Trustee shall have no responsibility with respect to these matters, except on written direction by the beneficiary(ies), their appointed beneficiary to hold the power of direction, or their third-party Director appointed pursuant to Article Three, and only after payment to the Trustee of all funds necessary to carry out such direction.'),
    pb(),
    p('\t\t\t\t2.4(d)\tRight of Assignment.  The beneficiaries, or any one of them, shall have the full right to assign their interest in the trust in writing to any other person or entity, subject to any written agreement between or among beneficiaries that may restrict such assignment. Such restrictive agreement must be filed with the Trustee in order to be binding on the Trustee.  No assignments or other instruments conveying the interest of any beneficiary hereunder shall be binding on the Trustee or any beneficiary hereunder until the original or an executed duplicate copy of such assignment or other instrument is filed with and its receipt is acknowledged by the Trustee.'),
    pb(),
    p('\t\t\t\t2.5\tRight of Succession of Beneficiaries.  Upon the death or dissolution of a beneficiary during the term of this trust, his, her, or its beneficial interest in this trust shall be vested in the following persons or entities who shall succeed to the deceased or dissolved beneficiary\'s interest (listed in the same order as the beneficiary is listed in paragraph.2.2).'),
    pb(),
    p('Name and Address of Successor in Interest to Beneficiary:'),
    pb(),
    p('1.   Successor Beneficiary named by first beneficiary listed in Paragraph 2.2 (in order of priority, not jointly)'),
    pb(),
    p(`\t\t1st Successor: ${d.successor1}`),
    pb(),
    p(`\t\t2nd Successor: ${d.successor2}`),
    pb(),
    p('2. Successor Beneficiary named by second beneficiary listed in Paragraph 2.2 (in order of priority, not jointly)'),
    pb(),
    p('\t\t1st Successor:'),
    pb(),
    p('\t\t\t Successor Beneficiaries are not granted a present vested interest in this trust. The successor beneficiary\'s interest is a future interest only, which can be revoked by the primary beneficiary at any time prior to the death or dissolution of the primary beneficiary or the termination of the trust if it is created for a specified term of years.'),
    pb(),
    p('\t\tIf an individual beneficiary has not assigned his, her, or its interest during lifetime to another, and for any reason the successor named in this paragraph is unavailable for any reason to receive the interest of the beneficiary at the death of the beneficiary, that interest shall pass to the beneficiary\'s revocable living trust, and if none, to the estate of the named beneficiary, however, not to his/her heirs at law, unless the beneficiary dies intestate.'),
    pb(),
    p('\t\t\tIf the beneficiary is an entity that is merged, reorganized or dissolving or dissolved, the beneficial interest shall pass to the entity\'s successor in interest, and if none, to the entity\'s shareholders, members, partners, or owners of other interests in the entity, as the case may be.'),
    pb(),
    p('\t\t\t\tThe death or legal dissolution of any beneficiary of this trust shall not terminate this trust nor in any manner affect the powers of the Trustee of this trust.'),
    pb(),
    p('\t\t\t\t2.6\tLimitations on Rights and Powers of Beneficiaries.  Limitations on beneficiaries\' rights are as follows:'),
    pb(),
    p('\t\t\t\t2.6(a)\tNo Legal or Equitable Rights.  No beneficiary shall have any legal or equitable right, title, or interest in the real property interest held in this trust but shall have only the beneficial interest in this trust and the appurtenant personal property rights set forth in this instrument.'),
    pb(),
    p('\t\t\t\t2.6(b)\tNo Right to Trustee\'s Powers.   No beneficiary shall have the right to affect or carry out the powers of the Trustee.'),
    pb(),
    p('\t\t\t\t2.6(c)\tNo Right to Partition.  No beneficiary shall have the right to partition any real property, the title to which is held by the Trustee of this trust.'),
    pb(),
    p('\t\t\t\t2.6(d)\tRights to Contract.  The beneficiary(ies) has/have the right or authority to contract on behalf of the trust or for or in the name of the Trustee or to bind the Trustee in any way through the beneficiary\'s action.  The beneficiary(ies) are not the agents of the Trustee, and the Trustee is not the agent of the beneficiary(ies).'),
    pb(),
    p('\t\t\t\t2.6(e)\tNo Right to Use Trustee Name.  The name of the Trustee shall not be used by the beneficiaries in connection with any business transactions or publicity without the written consent of the Trustee.'),
    pb(),
    pCenterBold('ARTICLE THREE'),
    pCenterBold('DIRECTORS'),
    pb(),
    p('\t\t\t\t3.1\tAppointment of Director.  The power of direction over the actions of the Trustee that is held by the beneficiary(ies) shall be deemed a fully assignable personal property interest under the terms of this trust.  Pursuant to Paragraph 2.4(a), the beneficiary or beneficiaries may, but shall not be required to, unbundle this interest from other rights held by the beneficiary(ies) and transfer his/her/its power of direction to a person who is not a beneficiary of this trust.  The person to whom the power of direction is transferred shall be known as the Director.  The assignment of the power of direction to another shall not affect any other rights provided to the beneficiary(ies) in this trust.  If the appointment by the beneficiary(ies) of a Director is not made in this Paragraph 3.1 upon the execution of this Trust, the appointment shall be made in a separate writing delivered to the Trustee and appended to the trust instrument and made part hereof.'),
    pb(),
    p(`\t\t\t\tThe initial beneficiary of the trust hereby appoints as the initial Director ${d.director}`),
    pb(),
    p('\t\t\t\t3.2\tRemoval and Replacement of Director.'),
    pb(),
    p('\t\t\t\t3.2(a)\tRemoval.  The beneficiary(ies) or the majority of them shall have the right to remove a Director appointed pursuant to this Article Three at any time with or without cause and such removal shall take effect immediately upon written notice to the Director and the Trustee.  If there are multiple beneficiaries, such removal shall be done only by majority vote and the beneficiaries shall immediately notify the Trustee in a writing signed by the majority of beneficiaries.'),
    pb(),
    p('\t\t\t\t3.2(b)\tReplacement.  Such notification shall indicate whether the beneficiary(ies) shall then retain the power of direction unto himself/herself/themselves or whether a new Director has been appointed.  If a new Director is appointed, such notification to the Trustee must include the name, address, phone number, and email address of the new Director.  Until the Trustee receives such written notification, the Trustee shall be held harmless by the beneficiary(ies) for actions taken under the direction of the replaced Director.'),
    pb(),
    p('\t\t\t\t3.3\tResignation of Director.  The Director shall have the right to resign by giving four days written notice to the beneficiaries and the Trustee.'),
    pb(),
    p('\t\t3.4\tAutomatic Termination of Director.  In the event the Director is compelled to act in its capacity by citation, court order, or generally against its will, or under duress, its power of direction shall automatically terminate immediately prior to such act.'),
    pb(),
    p('\t3.5\tSuccessor Directors by Default.  In the event of the death, dissolution, or removal of the Director where no successor Director is named by the beneficiary(ies) and the beneficiary(ies) have failed to notify the Trustee that he/she/they intend[s] to retain the power of direction unto Himself, the Trustor directs that the following shall serve as successor Director in the following order of priority:'),
    pb(),
    p('Name and Address:'),
    pb(),
    p(`-  \t${d.successorDirector1}`),
    pb(),
    p(`\t- ${d.successorDirector2}`),
    pb(),
    pCenterBold('ARTICLE FOUR'),
    pCenterBold('TRUSTEES'),
    pb(),
    p('\t\t\t\t4.1\tDuties and Responsibilities.  Legal and equitable title to the real property held by this trust shall be held solely by the Trustee.  The Trustee shall have the following duties, in addition to any other duties expressed in other paragraphs of this agreement:'),
    pb(),
    p('\t\t\t\t4.1(a)\tActing Only When Authorized.  While the Trustee is the sole legal owner of the real property held by this trust, it is understood and agreed by the Trustee and all current and future beneficiaries of this trust that the Trustee shall deal with the real property only when authorized to do so in writing by the beneficiary(ies), by a beneficiary selected by multiple beneficiaries to hold the power of direction, or by the Director who is properly appointed in accordance with Paragraph 3.1.'),
    pb(),
    p('\t\t\t\t4.1(b)\tTrustee to Fulfill Direction.  Unless directed in writing by the beneficiary(ies) or by the Director, the Trustee has no power to control or influence the real property or any use thereof, shall have no duty to maintain the trust property or make it productive, and shall have no duty to collect any proceeds or make any payments on behalf of the trust.  While the Trustee is the sole title holder of the real property held in this trust, the beneficiary(ies), or the Director if one is appointed, has the sole right to direct the action of the Trustee.'),
    pb(),
    p('4.1(c)\tDuty to Sell or Exchange.  Upon written direction of the beneficiary(ies) or the Director if one is appointed, the Trustee shall sell, transfer, or convey the property as directed. Upon written direction of the beneficiary(ies) or the Director, the Trustee shall deliver the property of the trust to the then-current beneficiaries in the shares set forth in paragraph 2.2 or their successors in interest as provided in paragraph 2.5, or to whom they shall designate in writing. Trustee will execute and deliver deeds, mortgages, or any other documents when so directed by the holder of the power of direction.'),
    pb(),
    p('\t\t\t\t4.1(d)\tNotification of Claims.  In the event the Trustee receives notice of claims or action against the trust, the Trustee shall promptly notify the beneficiary(ies) at their most recent address of record or by email.  The beneficiary(ies) has/have the duty to respond to such claims or notices and the Trustee shall be held harmless from further action to respond to such claims or notices.'),
    pb(),
    p('\t\t\t\t4.2\tTrustee Powers.  Solely with the consent and written direction of the beneficiary(ies) or their appointed Director, and subject to any limitations stated elsewhere in this instrument, to carry out the purposes of the trust created under this instrument, the Trustee shall have the following powers:'),
    pb(),
    p('4.2(a)\tTo hold the legal and equitable title to all of the trust property, and to do all things and perform all acts necessary and proper for the protection of the trust property and for the interest of the beneficiary(ies) in the property of the Trust, subject to the restrictions, terms, and conditions set forth herein and in any deed conveying an interest in real property to this trust;'),
    pb(),
    p('4.2(b)\tTo purchase or sell any real property, to assume mortgages upon the property for the Trust at such times and on such terms as directed by the beneficiary(ies) or the appointed Director; to execute notes, deeds, contracts, options, mortgages upon such terms as directed by the beneficiary(ies) or the Director, and otherwise deal with the trust property or its proceeds from disposition as may be directed by the beneficiaries or the Director;'),
    pb(),
    p('4.2(c)\tTo retain property received into the trust at its inception or later added to the trust, without regard to whether the trust investments are diversified.'),
    pb(),
    p('\t\t4.3\tPower to Delegate.  The Trustee may from time-to-time delegate in writing to a special trustee, or an agent acting on its behalf, the Trustee\'s authority to take actions hereunder, with the written direction of the Beneficiary(ies) or Director, if any.'),
    pb(),
    p('\t\t\t\t\t\t\t\t4.4 \tLimitations on Trustee\'s Powers.  The Trustee shall not have or exercise any powers other than those expressly granted to it under the terms of this Agreement and is not authorized to engage in any activity not necessary to the stated purposes of the trust.  The Trustee shall not transact business within the meaning of applicable state law, or any other law.'),
    pb(),
    p('\t\t\t\t\t\t\t4.5\tRemoval; Resignation; Replacement of Trustee.'),
    pb(),
    p('\t\t\t\t\t\t\t\t4.5(a)\tRemoval and Replacement. The beneficiary(ies) shall have the right to remove and replace the Trustee at any time by written notice to the Trustee.  If the right to remove and replace the Trustee is held by multiple beneficiaries, the removal and the replacement of a Trustee shall be by written majority vote. Such removal shall become effective immediately. Upon written notification of removal, the acting Trustee shall immediately forward by overnight delivery service all trust records in his/her/its possession to the successor Trustee, or if none yet appointed, to the beneficiary(ies). The beneficiary(ies) shall replace the removed Trustee promptly with a successor Trustee who shall accept his/her/its appointment in writing.'),
    pb(),
    p('\t\t\t\t\t\t\t4.5(b)\tResignation and Replacement.  The Trustee may at any time resign as Trustee hereunder by hand delivery or by mailing by registered mail (to the address last known by the Trustee) a copy of its written resignation to each of the then beneficiaries hereunder and to any Director then acting pursuant to Paragraph 3.1 at the addresses last known to the Trustee.  Such resignation shall become effective upon the appointment of a Successor Trustee as provided in Paragraph 4.5(a) above, or without further notice thirty (30) days from the date of such mailing or delivery, whichever is earlier.'),
    pb(),
    p('\t\t\t\t\t\t\t\t4.5(c)\tDuties of Terminated Trustee. The Trustee shall immediately convey the Trust property to the successor Trustee. Thereafter, the terminated Trustee shall still be bound under law by the duty of loyalty and is therefore enjoined not to disclose any of the affairs of the trust, the identities of any of the beneficiaries, past or present, or the location and/or description of trust assets except under a lawful court order issued by a duly constituted court of competent jurisdiction in the legal domicile of the trust. Failure to maintain trust confidentiality shall make the Trustee personally liable for any damages, which might ensue or be sustained by any party whatsoever or whomever as a result of said breach of trust (refer also to paragraph 5.1 regarding limitations of Trustees).'),
    pb(),
    p('\t\t\t\t\t\t\t\t4.5(d)\tTrustee Lien.  A departing or departed Trustee shall continue to have a lien on trust property if amounts due the Trustee remain unpaid; the provisions of paragraph 5.3 herein shall continue to apply.'),
    pb(),
    p('\t\t\t\t\t\t\t4.5(e)\tSuccessor Trustee Rights and Duties.  Upon accepting the office of Trustee, the successor Trustee shall succeed to the title of all property held by this trust and shall be subject to all provisions of this Agreement.  A successor Trustee properly appointed under the terms of this Article Four who has accepted the role of successor Trustee in writing shall be vested with all the estate, rights, powers, trusts, duties, and obligations of his/her/its predecessor. All rights, powers, authority, immunity, and discretion herein granted or conferred upon the original Trustee shall survive to and may be exercised or applied in the same manner and to the same extent by or for any successor or substitute that may at any time be acting hereunder.  Each Trustee shall be responsible only for its own acts or omissions.  A successor Trustee shall not be required to audit or investigate the acts or administration of any predecessor and shall be relieved of all liability for failing to do so. Reference herein to the Trustee shall include any company, corporation, or association that may become a successor trustee or any successor corporation which may succeed to a corporate trustee\'s business and such successor shall be bound by all the terms of this Agreement.'),
    pb(),
    p('\t\t\t\t\t\t\t\t4.5(f)\tDefault Successor Trustees.  In the case of the resignation, refusal, removal, or inability to act of a Trustee where no successor has been effectively appointed under this Article Four, the following person(s) or entities shall serve as Trustee:'),
    pb(),
    p(`\t\t\t\t\t\t\t1: ${d.defaultSuccessorTrustee1}`),
    pb(),
    p(`\t\t\t\t\t\t\t2: ${d.defaultSuccessorTrustee2}`),
    pb(),
    p('\t\t\t\t\t\t\t\t\tThese named persons or entities shall act successively in this order'),
    pb(),
    p('\t\t\t\t\t\t\t4.5(g)\tFailure to Appoint Successor Trustee.  If no successor is duly appointed under 4.5(a) or within the thirty (30) days provided in 4.5(b), and no successor Trustee named in 4.5(e) is available to act, the Trustee may convey the Trust property by Quit Claim or Trustee\'s Deed, to the then beneficiaries in accordance with their respective interests hereunder. Upon the delivery of said Quit Claim or Trustee\'s Deed to the Recorder of Deeds or the Registrar of Title for recordation, the Trust hereby created shall be terminated.'),
    pb(),
    p('\t\t\t\t4.6\tDuty Not to Disclose. The Trustee is bound never to reveal the name of the beneficiary nor the beneficiary\'s location nor ever to allow anyone to view the trust agreement without the unanimous written consent of the owners of the beneficial interests or a written court order (issued by a duly constituted court of competent jurisdiction in the legal domicile of the trust), which court order contains an indemnification clause protecting the Trustee from a lawsuit for breach of trust. The Trustee\'s only recourse if this provision of paragraph 4.6 is not adequately met is to resign immediately without taking any action other than forwarding all documents to the new Trustee in accordance with paragraph 4.5.'),
    pb(),
    pCenterBold('ARTICLE FIVE'),
    pCenterBold('TRUSTEE LIMITATIONS'),
    pb(),
    p('\t5.1\tNo Personal Obligation; No Warranty by Trustee. The Trustee shall not have or be required to enter into any personal obligation or liability in acting in accordance with the power of direction and shall not be required to make any conveyances or other instruments affecting real estate owned by the trust while money is owed to it as Trustee.  It is also expressly understood and agreed by and between the parties hereto, that any undertaking, action, or signature on the part of the Trustee shall not constitute a personal warranty, indemnity, representation, covenant, undertaking, or agreement, nor shall it be for the purpose or with the intention of binding the Trustee personally.  Such Trustee actions shall be made and are intended solely in accordance with this instrument and are executed and delivered by said Trustee solely in the exercise of the powers conferred upon the Trustee and not in his/her/its personal capacity.  No personal liability or personal responsibility is assumed by, nor shall at any time be asserted or enforceable against, the Trustee on account of any warranty, indemnity, representation, covenant, undertaking, action, or agreement of said Trustee undertaken pursuant to this instrument or its related instruments, either express or implied, and all such personal liability, if any, is expressly waived and released.'),
    pb(),
    p('5.2\tLimitation of Liability.  The Trustee shall not be required to obligate itself individually or be liable to pay or incur the payment of any damages, attorneys\' fees, fines, penalties, forfeitures, costs, charges, or other sums of money whatsoever for acting on behalf of this trust and fulfilling its duties as Trustee.  The Trustee shall have no individual liability arising from its holding of title to property held in this trust.  Further, the Trustee shall have no liability with respect to acts done, contracts entered into or indebtedness incurred with respect to the trust estate in fulfilling its duties under this Agreement.  Liability for any Trustee actions shall not exceed the assets of the trust estate and only trust estate assets shall fulfill the payment and discharge of any such liability or obligation.'),
    pb(),
    p('\t\t\t\t\t5.3\tReimbursement and Indemnification of Trustee.'),
    pb(),
    p('\t5.3(a)\t Reimbursement on Demand. The beneficiary or beneficiaries, jointly and severally, agree that on demand they will pay to the Trustee amounts to reimburse the Trustee for payments made or liabilities incurred by the Trustee, including breach of contract (unless fraud is committed by Trustee), injury to person or property, fines or penalties under any law, or otherwise, and money paid as a result of being made a party to any litigation deriving from holding title to the trust property together with its expenses, including reasonable attorneys\' fees.'),
    pb(),
    p('\t5.3(b)\t Indemnification. The beneficiary(ies) indemnify and hold the Trustee harmless from any and all payments made, or liabilities incurred by the Trustee by reason of serving as Trustee under this Agreement. All amounts paid or incurred by the Trustee, as well as Trustee compensation under this Agreement, shall constitute a first priority lien on the property held by this trust while due and owing but unpaid by the beneficiary(ies).'),
    pb(),
    p('\t5.3(c)\tTrustee Advances.  In case the Trustee shall make any advances of money or incur any liability or be obliged to pay out any money, including attorney\'s fees, by reason of its being Trustee hereunder, the beneficiaries do hereby agree, jointly and severally, on behalf of themselves, their heirs, executors, administrators and assigns, to pay any and all of such advancements, disbursements or liabilities, including attorney\'s fees, on demand by Trustee with interest on any unpaid amounts at 10 percent (10%) per annum.  In case of nonpayment within sixty (60) days after demand, the Trustee shall have the right and is hereby authorized and directed to sell sufficient trust property to pay the debt or advancement, and after deducting its own reasonable compensation, to pay the balance thereof to the beneficiaries as their interest is listed in paragraph 2.2.  However, the Trustee shall not be required to make any advances or to prosecute or defend any legal proceeding involving the Trust.  The Trustee shall not be obligated to incur any liability for the prosecution or the defense of any legal proceeding unless it shall be furnished with funds sufficient to meet such liability or be indemnified to its satisfaction in respect thereto.'),
    pb(),
    p('\t\t\t\t\t5.4\tCharging Order.  If a Charging Order shall be in effect at the time of disbursement to the beneficiaries, then all funds after reimbursement to the Trustee shall be accumulated and not disbursed.'),
    pb(),
    p('\t\t\t\t\t5.5\tLimitation on Power to Bind Beneficiaries.  The Trustee shall not have any power to bind the beneficiary(ies) personally.  Any person contracting with the Trustee shall look solely to the assets of the trust for payment under such contract, or for the payment of any debt, mortgage, judgment, or decree, or for any money that may otherwise become due or payable, whether by reason of failure of the Trustee to perform the contract, or for any other reason, and neither the Trustee nor the beneficiary(ies) shall be liable personally therefor.'),
    pb(),
    p('\t\t\t\t\t5.6\tAssignments and Conveyances of Interests.  No assignments or other instruments conveying the interest of any beneficiary hereunder shall be binding on the Trustee in any manner until the original of such assignment or other instrument, in such form as the Trustee may approve, is deposited with the Trustee, and accepted in writing by said Trustee. Furthermore, any assignments or conveyances of the beneficial interest(s), or any portion of a beneficial interest hereunder shall be conveyed subject to the existing conditions of the power of direction and shall in no way alter, amend, or revoke the powers of the direction.'),
    pb(),
    pCenterBold('ARTICLE SIX'),
    pCenterBold('DEALINGS WITH THIRD PARTIES'),
    pb(),
    p('\t\t\t\t\t\t\t\t6.1\tNo Disclosure by Trustee or Beneficiaries. The Trustee shall not release information regarding this trust or its beneficiaries.  A beneficiary shall not disclose the identity of any other beneficiary.'),
    pb(),
    p('\t\t\t\t\t\t\t\t6.2\tThird-Party Dealings with Trustee.  No person dealing with the Trustee with respect to trust property shall be obliged to see to the application of any purchase money, rent or money borrowed or otherwise advanced on the property or to see that the terms of this Trust Agreement have been complied with.  No person shall inquire into the authority, necessity, or expediency of any act of the Trustee or be privileged to inquire into any of the terms of this Trust Agreement.'),
    pb(),
    p('\t\t\t\t\t\t\t\tAn instrument executed by a then duly and properly acting Trustee shall be conclusive evidence in favor of a person claiming any right, title, or interest under the Trust that, at the time of delivery of such instruments, the Trust created under this Agreement was in full force and effect and that the instrument was executed in accordance with the terms and conditions of this Agreement and all its amendments.  If a conveyance is made to a duly appointed and acting successor trustee, it shall be deemed that the Successor Trustee is vested fully with all the title, estate, rights, powers, duties, and obligations of the predecessor Trustee.'),
    pb(),
    pCenterBold('ARTICLE SEVEN'),
    pCenterBold('TRUSTEE COMPENSATION'),
    pb(),
    p('\t7.1\tReasonable Compensation.  The beneficiary(ies) jointly and severally agree(s) to pay reasonable fees to the Trustee for his/her/its services.'),
    pb(),
    p('7.1(a)\tAmount of Compensation. Such amounts shall not exceed $1.00 per year.'),
    pb(),
    p('7.1(b)\tCompensation for Additional Services.  Furthermore, the Trustee shall receive reasonable compensation for making deeds or other instruments, performing additional services, or retaining attorneys, accountants, or agents.'),
    pb(),
    p('\t7.1(c)\tTrustee\'s Lien. The beneficiaries hereunder jointly and severally agree to pay the aforesaid fees hereunder and the Trustee shall have a priority lien on all trust assets, therefore.  No instruction by the beneficiaries nor agreements made by the Trustee hereunder shall be binding so long as there remain any unpaid claims for the Trustee\'s compensation, nor may the Trustee be compelled to perform any of his duties so long as this condition prevails.'),
    pb(),
    p('\t\t\t\t\t\t\t\t7.2\tReimbursement of Expenses. The Trustee may be reimbursed for reasonable expenses incurred on behalf of the trust in the conduct of trust affairs.'),
    pb(),
    pCenterBold('ARTICLE EIGHT'),
    pCenterBold('NOTICES'),
    pb(),
    p('\t\t\t\t\t\t\t\t8.1\tMailing of Notices.  The Trustee shall mail any notices to the last known address of the current beneficiary or beneficiaries of the trust.  The Trustee shall not be responsible for a beneficiary\'s failure to notify the Trustee of a change of address.'),
    pb(),
    p('\t\t\t\t\t\t\t\t8.2.  Other Delivery Forms. Except in the case of legal notices, the Trustee may also notify the beneficiary(ies) by telephone, text, facsimile, email, or other form of instant messaging of important matters regarding the trust or trust property.'),
    pb(),
    pCenterBold('ARTICLE NINE'),
    pCenterBold('ADMINISTRATIVE PROVISIONS'),
    pb(),
    p('\t\t\t\t\t\t\t\t9.1\tNo Recording.  This Trust Agreement shall not be placed of record in the Recorder\'s Office in the county in which the property is situated or elsewhere and the recording of the same shall not be considered as notice of the rights of any person hereunder derogatory to the title of the Trustee, nor shall any person dealing with the Trustee be privileged or required to inquire into the necessity or expediency of any act of the Trustee or of the provisions of this instrument.'),
    pb(),
    p('\t\t\t\t\t\t\t\t9.2 \tTermination of Trust'),
    pb(),
    p('\t\t\t\t\t\t\t\t9.2(a)\tBy Trustee.  The Trustee shall have no right to terminate this Agreement.'),
    pb(),
    p('\t\t\t\t\t\t\t\t9.2(b)\tThis trust shall continue for: Indefinitely, or until the trust, property in the trust is transferred or sold. If the trust terminates after a specified term of years, the Trustee shall distribute all trust property to the then beneficiaries in their proportionate shares or percentages of record.'),
    pb(),
    p('\t\t\t\t\t\t\t\t9.2(c)\tBy Beneficiaries.  This trust may be terminated at any time by the beneficiary(ies) with written notice of termination signed by all beneficiaries and delivered to the Trustee. The Trustee shall then within ten (10) days execute any and all documents necessary to vest fee simple marketable title to any and all property titled in the name of the Trustee to the beneficiary(ies), subject to the provisions of paragraphs 5.3(c) and 7.1.'),
    pb(),
    p('\t\t\t\t\t\t\t\t9.3\tAmendments.  This Trust Agreement contains the entire understanding among the parties and may be amended, modified, or revoked by the beneficiary, or by the majority of all beneficiaries, in whole or in part only by written instrument signed by the beneficiary or all of the beneficiaries and delivered to the Trustee.  An exercise of the power of amendment substantially affecting the duties, rights, and liabilities of the Trustee shall be effective only if agreed to in writing by the acting Trustee.'),
    pb(),
    p('\t\t\t\t\t\t\t\t9.4\tMiscellaneous.'),
    pb(),
    p('\t\t\t\t\t\t\t\t9.4(a)\tGender reference. Any reference to one gender shall also refer to the other.'),
    pb(),
    p('9.4(b)\tBinding on Heirs.  The terms and conditions of this Agreement shall inure to the benefit of and be binding upon any successor trustee, as well as upon the executors, administrators, heirs, assigns, and all other successors in interest of the beneficiaries.'),
    pb(),
    p('\t\t\t\t\t\t\t\t9.4(c)\tTitles.  Titles of articles and paragraphs of this Agreement are for convenience only and shall not be used to construe the meaning of any provision of this Agreement.'),
    pb(),
    p('\t\t\t\t\t\t\t\t9.4(d)\tNo Bond.  No bond or other security shall be required of any Trustee for the faithful performance of its duties in such capacity.'),
    pb(),
    p('\t\t\t\t\t\t\t\t9.5\tGoverning Law.  The Trustee shall take title to property under the provisions of the laws of Illinois and hereby agrees to hold the said real property for the uses and purposes and upon the terms herein set forth.  Where contradictions, conflicts, or ambiguities occur between this Trust Agreement and cited laws or statutes, the text herein shall prevail.'),
    pb(),
    p('\t\t\t\t9.6\tSpendthrift Provision.  No income or principal distributable or to become distributable with respect to the trust shall be transferable, assignable, or subject to being in any manner whatsoever anticipated, charged or encumbered by any person beneficially interested in the trust, or subject to interference or control by any creditors of said person, or subject to any claim for alimony or the support of a spouse pursuant to a decree of separate maintenance or separation agreement, or to being taken or reached by any legal or equitable process in satisfaction of any debt, liability or obligation of said person prior to its receipt by said person; provided, however, that the provisions of this paragraph shall not prevent (i) the exercise of, or transfer of income or principal pursuant to the exercise of, any power to assign granted in paragraphs 2.4(d) and 3.1 or (ii) the exercise of any power to disclaim or renounce the whole or any part of an interest in this trust.'),
    pb(),
    p('\t\t\t  \tThe beneficiary shall have no other right to alienate, encumber or hypothecate its interest in the trust estate, nor shall such interest be subject to claims of the beneficiaries\' creditors or be liable for attachment, execution, or other process of law. The interest of each beneficiary shall be free from the control or interference of any creditor of a beneficiary or any spouse of a married beneficiary. Such interest shall not be subject to attachment or susceptible to anticipation or alienation. This paragraph shall not be construed as restricting in any way the exercise of any powers or discretions.'),
    pb(),
    p('\t\t\t9.7\tSeverability.  The invalidity of any portion of the trust agreement will not be deemed to affect the validity of the other portion of this agreement. In the event any provision of this agreement shall be held to be invalid by a court of competent jurisdiction, the parties agree that the remaining portions shall remain in full force and effect, as if the agreement had been executed by the parties subsequent to the expungement of the invalid provision(s). All parties to this contract agree to correct scrivener\'s errors should they arise.'),
    pb(),
    p(`IN WITNESS WHEREOF and in consideration of the foregoing, the Trustee accepts the trust thus committed to it and agrees to act in accordance with the terms stated above. The Trustee and the beneficiary(ies) hereby execute this Agreement on ${fmtDateShort(d.trustDate)}.`),
    pb(),
    pBold(`BENEFICIARY: ${d.beneficiary}`),
    pb(),
    new Paragraph({
      spacing: { after: 160 },
      children: [
        new TextRun({ text: 'Signature: ______________________________________________ Date: _________________', size: 24, font: 'Times New Roman' }),
      ]
    }),
    pb(),
    p(`Beneficiary: ${d.beneficiary} Member Manager: ${d.director}`),
    pb(),
    pBold('NOTARY ACKNOWLEDGEMENT:'),
    pb(),
    p('State of: ____________________________   County of:_________________________________'),
    pb(),
    p(`On ____________________, before me, personally appeared ${d.director},`),
    p('Insert date'),
    pb(),
    p('who proved to me on the basis of satisfactory evidence to be the person whose name is subscribed to the within instrument and acknowledged to me that he/she executed the same in his/her authorized capacity, and that by his/her signature on the instrument the person, or the entity upon behalf of which the person acted, executed the instrument.'),
    pb(),
    p('I certify under penalty of perjury under the laws of the State of _____________________ that the foregoing paragraph is true and correct.'),
    pb(),
    p('WITNESS my hand and official seal.'),
    pb(),
    p('Signature _________________________________'),
    pb(),
    new Paragraph({
      spacing: { after: 160 },
      children: [
        new TextRun({ text: 'Notary Acknowledgement', size: 24, font: 'Times New Roman', bold: true }),
        new TextRun({ text: '                                                    (Seal)', size: 24, font: 'Times New Roman', italics: true }),
      ]
    }),
    pb(),
    pCenterBold('SCHEDULE OF TRUST PROPERTY'),
    pb(),
    p(`${d.propAddress}  Property tax I.D. #${d.pin}`),
    pb(),
    p(`Legal description: ${d.legalDesc}`),
    pb(),
    p('Accepted by Trustee'),
    pb(),
    new Paragraph({
      spacing: { after: 60 },
      children: [
        new TextRun({ text: `Dated ${fmtDateShort(d.trustDate)}`, size: 24, font: 'Times New Roman', bold: true }),
        new TextRun({ text: `\t${d.trustee}, Trustee, By: ${d.director},`, size: 24, font: 'Times New Roman', bold: true }),
      ]
    }),
    pBold('          Its: Member/Manager'),
    pb(),
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: '________________________________________\t\t\t\t__________________', size: 24, font: 'Times New Roman' })]
    }),
    p('Signature\tManager / Member ' + d.trustee + '\t\t\t  Date'),
    pb(),
    p(`\t\t${d.trustName}               \t                           1                               ${fmtDateShort(d.trustDate)}`),
  ];
  return new Document({ sections: [{ properties: PAGE_PROPS, children }] });
}

function buildCertification(d) {
  const children = [
    new Paragraph({
      spacing: { after: 160 },
      children: [new TextRun({ text: 'CERTIFICATION OF LIVING TRUST', size: 24, font: 'Times New Roman', bold: true, italics: true })]
    }),
    pb(),
    p('STATE / COMMONWEALTH OF ___________, COUNTY OF ______________, SS:'),
    pb(),
    p('The undersigned, being first duly sworn (or affirmed) states as follows:'),
    pb(),
    new Paragraph({
      spacing: { after: 160 },
      children: [
        new TextRun({ text: 'Trustee. The undersigned is the trustee of the ', size: 24, font: 'Times New Roman' }),
        new TextRun({ text: `${d.trustName}`, size: 24, font: 'Times New Roman', bold: true }),
        new TextRun({ text: ' (the "', size: 24, font: 'Times New Roman' }),
        new TextRun({ text: 'Trust', size: 24, font: 'Times New Roman', bold: true }),
        new TextRun({ text: '"). The Trust (i) is not subject to administration by any probate court or similar court system, (ii) continues to be in full force and effect, and (iii) has not been revoked.', size: 24, font: 'Times New Roman' }),
      ]
    }),
    pb(),
    p('Trust Provisions. Attached to this certification are true and accurate copies of the specific provisions of the Trust that evidence the following.'),
    pb(),
    p('- The establishment of the Trust.'),
    p('- The identity of the initial trustee(s).'),
    p('- The provisions regarding successor trustees.'),
    p('- The provisions reserving the rights to revoke the Trust and/or to change or amend its provisions.'),
    p('- The general administrative provisions.'),
    p('- The page(s) showing the signature(s) of the parties to the Trust agreement.'),
    pb(),
    p('The provisions that are not attached to this certification are personal. They include provisions regarding the distribution of assets and other private matters, and do not modify or otherwise affect the trustee powers.'),
    pb(),
    p('Certification and Agreement to Hold Harmless. The undersigned trustee certifies that the above statements are true and correct. All parties to whom this affidavit is given are entitled to rely on its accuracy. Such parties shall be held harmless by the undersigned and the successors of the undersigned.'),
    pb(),
    new Paragraph({
      spacing: { after: 60 },
      children: [
        new TextRun({ text: '_________________________________________    ', size: 24, font: 'Times New Roman' }),
        new TextRun({ text: '   _________________________', size: 24, font: 'Times New Roman' }),
      ]
    }),
    p(`Signature of Trustee: Member / Manager of                  DATE`),
    p(`${d.trustee} – ${d.director}`),
    pb(),
    p('Subscribed and sworn to (or affirmed) before me this ____ day of ____________________, 20______.'),
    pb(),
    p('______________________________________'),
    p('Signature - Notary Public'),
  ];
  return new Document({ sections: [{ properties: PAGE_PROPS, children }] });
}

function buildAppointment(d) {
  const children = [
    pCenterBold('APPOINTMENT OF TRUSTEE', { size: 26 }),
    pb(),
    p(`The UNDERSIGNED Beneficiary(ies) represent that they are one hundred percent of the beneficiaries to that certain trust agreement, dated on the ${ordDay(d.trustDate)} of ${monthName(d.trustDate)} ${year(d.trustDate)}, and which is known as ${d.trustName} and they hereby appoint ${d.trustee}, whose address is ${d.trusteeAddr} to represent all of their beneficial interests in said trust as their Trustee for said interests jointly and severally in accepting written direction from the below named beneficiaries and their successors and serving the said trust as trustee, following their written direction in the performance of such ministerial tasks as are authorized and required by the above referenced trust agreement under which said trust has been created.`),
    pb(),
    p('In the event of the death, disability, incapacity or refusal to act of the Trustee appointed herein, and no successor has been appointed, then the duties of the Trustee shall devolve upon the below beneficiaries and on their assigns as such time, (s)he shall transfer all trust assets by Trustee Deed to said Trustee upon duly authorized direction of the beneficiaries, and any further actions taken by the above Trustee shall be personal, and not as the authorized or as the lawful holder of the Trustee powers over the aforesaid trust.  So say we all, holders of 100% of the beneficial interests.'),
    pb(),
    p(`${d.beneficiary} Beneficiary of ${d.benefitPct} undivided interest`),
    pb(),
    pb(),
    pBold('ACCEPTANCE/RESIGNATION OF TRUSTEE'),
    pb(),
    p(`TO: The above beneficiaries of the above said trust.  This is to advise you that as of the ${ordDay(d.trustDate)} of ${monthName(d.trustDate)}, ${year(d.trustDate)}, I hereby accept the position of Trustee subject to the power of direction over the aforesaid trust together with my fiduciary to each and all of you.  Please govern yourselves accordingly.`),
    pb(),
    pb(),
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: '\t\t\t\t\t\t\t________________________________', size: 24, font: 'Times New Roman' })]
    }),
    new Paragraph({
      spacing: { after: 240 },
      children: [new TextRun({ text: '\t\t\t\t\t\t\tTrustee', size: 24, font: 'Times New Roman' })]
    }),
    pb(),
    new Paragraph({
      spacing: { after: 160 },
      children: [new TextRun({ text: 'THIS DOCUMENT IS NOT ', size: 24, font: 'Times New Roman', bold: true }),
                 new TextRun({ text: 'REQUIRED', size: 24, font: 'Times New Roman', bold: true, underline: { type: UnderlineType.SINGLE } }),
                 new TextRun({ text: ' TO BE RECORDED IN ANY COUNTY RECORDER\'S OFFICE', size: 24, font: 'Times New Roman', bold: true })]
    }),
  ];
  return new Document({ sections: [{ properties: PAGE_PROPS, children }] });
}

function buildSpecialWarrantyDeed(d) {
  const children = [
    p('THIS SPACE PROVIDED FOR RECORDER\'S USE ONLY:'),
    p('WHEN RECORDED RETURN TO:'),
    p(d.trustName),
    p(d.returnAddr1),
    p(d.returnAddr2),
    new Paragraph({
      spacing: { after: 160 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '000000', space: 1 } },
      children: [new TextRun({ text: '', size: 24, font: 'Times New Roman' })]
    }),
    pb(),
    pCenterBold('SPECIAL WARRANTY DEED', { size: 26 }),
    pb(),
    p('THE GRANTOR(S),'),
    p(`${d.grantor} ${d.grantorDesc}`),
    p(`for and in consideration of: ${d.consideration} grants, bargains, sells, conveys and specially warrants to the GRANTEE(S): ${d.trustName} the following described real estate, situated at ${d.propAddress}`),
    pb(),
    p(`County of: ${d.county}`),
    pb(),
    p('Legal Description:'),
    p(d.legalDesc),
    pb(),
    p(`Commonly known as: ${d.commonAddr} (the "Property")`),
    pb(),
    p('Subject to existing taxes, assessments, liens, encumbrances, covenants, conditions, restrictions, rights of way and easements of record, the Grantor hereby covenants with the Grantee(s) the following:'),
    pb(),
    p('1. Covenant Against Encumbrances: The Grantor guarantees that there are no encumbrances upon the property other than those that have been already disclosed to the Grantee.'),
    pb(),
    p('2. Covenant of Warranty: The Grantor guarantees to the Grantee and the Grantee\'s heirs, executors and administrators that they shall have full possession of the property and not be ejected from the premises.'),
    pb(),
    p(`Grantor hereby releases and waives all rights under and by virtue of the Homestead Exemption Laws of the State of ${d.state}.`),
    pb(),
    p(`Tax Parcel Number: ${d.pin}`),
    pb(),
    pCenter('Page 1 of 2', { size: 20 }),
    pb(),
    p('Grantor Signatures:'),
    pb(),
    ...sigPair(d.grantor),
    pb(),
    p('STATE OF ___________________, COUNTY OF ______________________, ss:'),
    pb(),
    p('Notary Acknowledgment:'),
    pb(),
    p('STATE OF __________________ )'),
    p('                            ) SS                              COUNTY'),
    p('OF ________________         )'),
    pb(),
    p('I, _________________________________ the undersigned, a Notary Public in and for said County'),
    p('                   Printed Notary Name', { size: 20 }),
    p(`and State aforesaid, DO HEREBY CERTIFY THAT ${d.grantor} are all personally known to me to be the same people whose names are subscribed to the foregoing instrument, as having executed the same, appeared before me this day in person and acknowledged that they signed, sealed, and delivered the said instrument as their free and voluntary act for the uses and purposes therein set forth including the release and waiver of the right of homestead given under by hand.`),
    pb(),
    p('_______________________________________________'),
    p('Notary Public Signature of person taking acknowledgment'),
    pb(),
    p('_________________________________'),
    p('Title (and Rank)'),
    pb(),
    p('My commission expires _____________'),
    pb(),
    pCenter('Seal          page 2 of 2', { size: 20 }),
  ];
  return new Document({ sections: [{ properties: PAGE_PROPS, children }] });
}

// ─── API ENDPOINT ─────────────────────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  try {
    const raw = req.body;
    const d = {
      trustDate:              raw.trustDate || '',
      trustName:              raw.trustName || '[TRUST NAME]',
      trustee:                raw.trustee || '[TRUSTEE]',
      trusteeAddr:            raw.trusteeAddr || '[TRUSTEE ADDRESS]',
      beneficiary:            raw.beneficiary || '[BENEFICIARY]',
      beneficiaryAddr:        raw.beneficiaryAddr || '',
      benefitPct:             raw.benefitPct || '100%',
      director:               raw.director || '[DIRECTOR]',
      successor1:             raw.successor1 || '[SUCCESSOR 1]',
      successor2:             raw.successor2 || '[SUCCESSOR 2]',
      successorDirector1:     raw.successorDirector1 || '[SUCCESSOR DIRECTOR 1]',
      successorDirector2:     raw.successorDirector2 || '[SUCCESSOR DIRECTOR 2]',
      defaultSuccessorTrustee1: raw.defaultSuccessorTrustee1 || '[DEFAULT SUCCESSOR TRUSTEE 1]',
      defaultSuccessorTrustee2: raw.defaultSuccessorTrustee2 || '[DEFAULT SUCCESSOR TRUSTEE 2]',
      propAddress:            raw.propAddress || '[PROPERTY ADDRESS]',
      county:                 raw.county || '[COUNTY]',
      state:                  raw.state || 'Florida',
      pin:                    raw.pin || '[PIN]',
      legalDesc:              raw.legalDesc || '[LEGAL DESCRIPTION]',
      commonAddr:             raw.commonAddr || raw.propAddress || '[PROPERTY ADDRESS]',
      grantor:                raw.grantor || '[GRANTOR]',
      grantorDesc:            raw.grantorDesc || 'a married couple, no homestead involved',
      consideration:          raw.consideration || 'One Dollar ($1.00) and other good and valuable consideration',
      saleDate:               raw.saleDate || '',
      returnAddr1:            raw.returnAddr1 || '7210 Manatee Ave #1278',
      returnAddr2:            raw.returnAddr2 || 'Bradenton, FL 34209',
    };

    const selected = raw.docs || [];
    const zip = new JSZip();

    const builders = {
      deed:      { label: 'Special_Warranty_Deed',      fn: buildSpecialWarrantyDeed },
      appt:      { label: 'Appointment_of_Trustee',     fn: buildAppointment },
      trust:     { label: 'Trust_Agreement',             fn: buildTrustAgreement },
      deedtrust: { label: 'Deed_to_Trustee',             fn: buildDeedToTrustee },
      cert:      { label: 'Certification_of_Living_Trust', fn: buildCertification },
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
