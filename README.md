# Tally

Personal finance, on the phone, backed by a Google Sheet you own.

This repository is the **shipped app only**, published to GitHub Pages at
https://jpasserin.github.io/tally/ — open it on the phone, then
Settings ▸ About ▸ Install. Each commit here is a snapshot of one released
version; the source, tooling and tests live in a private repository.

- `index.html` — the app, one file, no framework
- `backend.js` — the one file that touches the network: your own Apps Script
- `Code.gs` — that Apps Script, with placeholders for the token and sheet id
- `sw.js`, `manifest.webmanifest`, icons — the installable-app parts
