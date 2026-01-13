# Sample Outlook Add-in

A minimal Outlook task-pane add-in that inserts a quick reply template into a draft message. The add-in is fully static so it can be hosted on GitHub Pages.

## Setup
- URLs in `manifest.xml` point to `https://yksnit.github.io/atom-outlook/` (update if you rename the repo or host elsewhere).
- Update `ProviderName` and the reply template text in `taskpane.js` as you like.

## Run locally
- Serve the folder (so Office can load the files) with `python -m http.server 8000` from the repo root.
- Temporarily set the URLs in `manifest.xml` to `https://localhost:8000/...` while testing locally, then switch them back to the GitHub Pages URLs.

## Publish to GitHub Pages
- Push this repo to GitHub and enable Pages for the `main` branch (root or `/` folder).
- Your files will be served from `https://yksnit.github.io/atom-outlook/`.
- Confirm the manifest is reachable at `https://yksnit.github.io/atom-outlook/manifest.xml`.

## Sideload the add-in
- Outlook on the web: Settings (gear) → View all Outlook settings → Mail → Customize actions → Add-ins → Upload custom add-in. Provide the manifest URL from GitHub Pages or upload the local `manifest.xml`.
- Outlook desktop (new Outlook / M365): Get Add-ins → My add-ins → Add a custom add-in → Add from file, then pick `manifest.xml`.
- Once installed, open a new message and use the ribbon button “Insert reply” to open the task pane and drop in the template.

## Next steps
- Add additional permissions (e.g., calendar, messages) in `manifest.xml` when you are ready to expand scope.
- If you change the repo name or host, update all URLs and the `AppDomains` entry in the manifest.
