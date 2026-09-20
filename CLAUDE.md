# Land trust generator

Generates land trust documents with per state clause variations. Read `../DOC-GENERATOR-NOTES.md` for the shared document rules.

Node and Express, using `jszip` for packaging and `xlsx` for spreadsheet input.

## Run it
```
npm install
node server.js
```

## What lives where
- `server.js` the web server and the form
- `templates/` document templates
- `state-clauses.json` the per state clause text, which is where state differences live. Change clauses here rather than branching in code.
- `public/` static assets
- `test/` tests
- `DEPLOY.md` deployment notes

## Deploy
Railway, project named "Land Trust Doc Gen". Run `railway status` in this folder to confirm the link. Push to `main` to deploy.

## Related
`Trust-Generator`, `il-land-trust-docs`, and `landtrust-docprep` are separate apps in this same area. Check which one Marisa means before changing land trust output.
