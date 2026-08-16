# Deploying this update

Railway project **Land Trust Doc Gen** → service **Land-Trust** → deploys from
`mgotsch46/Land-Trust` branch `main`. Pushing to `main` triggers a deploy
automatically; nothing needs changing in Railway itself.

## Steps

```bash
cd /path/to/your/Land-Trust      # your local clone
git checkout -b template-fill    # optional but recommended
```

Copy everything from this folder over your clone, then:

```bash
git rm -r --cached . && git add -A     # picks up deletions + new binaries
git status                             # confirm the file list below
npm install
npm test                               # must print "All checks passed."
npm start                              # sanity check at localhost:8080
```

Then:

```bash
git commit -m "v2: template-fill engine for recording-exact output

- Replace docx-library reconstruction with .docx template filling
- Add templates/ with the four approved source documents
- Illinois statutory language retained; other states use generic clause set
- Remove Special Warranty Deed from the packet
- Add fidelity regression suite and /health endpoint"

git push -u origin main
```

Railway builds on push. Watch the deploy, then check:

- `https://land-trust-production.up.railway.app/health` → `{"ok":true,...}`
- Generate an Illinois packet and open the Deed to Trustee in Word

## Expected file changes

| Change | Path |
|---|---|
| rewritten | `server.js` (808 → ~250 lines) |
| rewritten | `public/index.html` |
| modified | `package.json` (drops `docx` dependency) |
| added | `templates/` (4 .docx files) |
| added | `state-clauses.json` |
| added | `test/verify.js` |
| added | `README.md`, `.gitignore` |
| delete | `download` (stray 27-byte file in repo root) |

## Rollback

```bash
git revert HEAD && git push
```

Or in Railway, redeploy the previous deployment from the service's Deployments tab.

## Verify before recording

Generate one packet and confirm in Word:

1. Page size 8.5 × 11, margins unchanged
2. Pagination matches your source documents
3. No `{{TOKEN}}` text anywhere
4. No yellow highlighting
5. Illinois packets still cite 765 ILCS 405/407/410/430/435
