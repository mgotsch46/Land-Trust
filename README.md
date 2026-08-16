# RenewEQ Land Trust Generator

Generates a four-document land trust closing packet as Word `.docx` files.

Live: https://land-trust-production.up.railway.app/

## Architecture — read this before changing anything

Documents are produced by **filling the original `.docx` templates** in `templates/`.
The generator unzips each template, substitutes `{{TOKEN}}` values inside `<w:t>`
text nodes, and rezips entry-for-entry. Every byte of page setup, styles, fonts,
headers, footers and embedded media is copied through untouched.

**Do not reintroduce the `docx` npm library to rebuild these documents.** v1.x
reconstructed each document from hand-coded paragraphs, which could not reproduce
the source layout exactly. County recorders reject documents whose margins,
pagination or spacing drift. Template-fill is the whole point of v2.

Verified by `npm test`:

- output XML structure is identical to the template except for text node contents
- `sectPr` (page size and margins) identical
- paragraph and run property counts identical
- all non-`document.xml` parts byte-identical
- no unresolved `{{TOKENS}}` leak into output

## Documents

| Key | File |
|---|---|
| `trust` | Trust Agreement |
| `appt` | Appointment of Trustee |
| `deedtrust` | Deed to Trustee |
| `cert` | Certification of Living Trust |

The Special Warranty Deed was removed in v2.0.

## Multi-state handling

Illinois is the default and retains all statutory citations
(765 ILCS 405 / 407 / 410 / 430 / 435 and the Property Tax Code exemption).

For any other state, `state-clauses.json` supplies a generic clause set: the
Illinois statutory references are replaced with state-neutral wording and the
state name is substituted throughout.

### Adding a state's transfer-tax exemption

The Deed to Trustee exemption block only appears when the state has a configured
citation. Add it in `server.js`:

```js
const TRANSFER_TAX_EXEMPTION = {
  Illinois: 'Section 31-45 of the Property Tax Code',
  Indiana:  '<citation confirmed with the county recorder>',
};
```

If a state has no entry, the exemption block is omitted entirely.

> The generic clause set is state-neutral wording, not legal advice. Have counsel
> confirm the language for any new state before recording.

## Fields

See `state-clauses.json` for clause tokens and `buildValues()` in `server.js`
for the full field map. Dates are submitted as `YYYY-MM-DD` and formatted per
document (long form, ordinal "12th day of March, 2026", or `MM/DD/YYYY`).

`STATE_UPPER` and `NOTARY_COUNTY_UPPER` are derived automatically.

## Local development

```
npm install
npm start          # http://localhost:8080
npm test           # fidelity regression suite
```

`GET /health` reports whether all four templates are present.

## Deployment

Railway project **Land Trust Doc Gen**, service **Land-Trust**, deploying from
`main`. Pushing to `main` triggers a deploy. The service listens on `PORT`
(default 8080).

## Updating a template

1. Edit the `.docx` in `templates/` directly in Word.
2. Keep `{{TOKEN}}` text intact — a token must stay inside a single run, so
   avoid retyping partially over one.
3. Run `npm test` before pushing.
