# WebEmbed for PowerPoint (Mac)

A PowerPoint **content add-in** that puts a live web view on a slide. You can embed:

- **Web pages**: any https site that allows embedding. YouTube and Vimeo links are converted to their players automatically.
- **HTML assets**: paste HTML/CSS/JS or open an `.html`/`.svg` file. It's stored *inside the .pptx*, so the asset doesn't need its own hosting. The add-in itself still needs to be reachable (see below).
- **Samples**: three ready-made animations (particle orbit, kinetic type, animated bar chart), copied into the deck when you pick one.

Each embed is a box you can move and resize like any shape. It's live in the slideshow, so animations play and videos and pages respond to clicks. You can put several embeds on a slide, each with its own content. In Normal view a bar under the content shows what's embedded, with **↻** (reload) and **Change**. The bar is hidden while presenting, and the content then fills the box.

## How it works

```
PowerPoint slide
 └─ content add-in frame  ← https://localhost:3000/embed.html (this project)
     └─ <iframe>          ← your URL, or your HTML via srcdoc (sandboxed)
```

Settings for each embed (URL or inline HTML, plus page width) are saved in the document through `Office.context.document.settings`.

## Setup (one time)

Requires Node 18+ and PowerPoint for Mac (Microsoft 365 / 2019+).

1. Trust a localhost HTTPS certificate. PowerPoint only loads add-ins over HTTPS. This adds a cert to your login keychain and asks for your password.

   ```bash
   npm run certs
   ```

2. Install the manifest into PowerPoint's developer add-ins folder.

   ```bash
   npm run sideload
   ```

## Every time you use it

1. Start the add-in server and leave it running.

   ```bash
   npm start
   ```

2. Quit PowerPoint fully (⌘Q) and reopen it the first time after sideloading.
3. Open **Home → Add-ins** (or **Insert → Add-ins → My Add-ins**), then choose **WebEmbed** under *Developer Add-ins*.
4. The box appears on the current slide. Pick a tab, paste a URL or HTML, and press **Embed**.

> **The server must be running when you present.** The add-in page is served from `localhost:3000`. If the server is down, the live embed won't load. PowerPoint falls back to a load error or the last snapshot image. For a deck you'll present elsewhere or share, host `public/` somewhere with HTTPS (see below).

## Presenting

- **Presenter View can't be clicked.** Your screen in Presenter View shows a still picture of the slide; the live page exists only on the audience screen. To use the page, move the pointer onto the audience screen and click there, or mirror your displays, or present without Presenter View. This is PowerPoint's behaviour ([office-js#947](https://github.com/OfficeDev/office-js/issues/947)) and an add-in can't change it.
- **Click once to focus.** The first click in the embed can go to giving it focus; if a button doesn't react, click it again.
- **Scrolling may change slides.** Some PowerPoint versions send the scroll wheel to slide navigation even over an add-in ([office-js#2662](https://github.com/OfficeDev/office-js/issues/2662)). Dragging the page's scrollbar still works.
- **A blank or frozen box** is PowerPoint showing a saved snapshot instead of the live page. Click PowerPoint's ↻ on the box (in Normal view) before you start the show.

## Hosted copy (GitHub Pages)

The add-in is published at **https://yuxuancao97.github.io/webembed/**. Every push to `main` rebuilds it from `public/` (`.github/workflows/pages.yml`), together with a `manifest.xml` that points at it. With the hosted copy you don't need `npm start` to present.

1. Install the manifest that loads from Pages, then quit and reopen PowerPoint.

   ```bash
   ADDIN_URL=https://yuxuancao97.github.io/webembed npm run sideload
   ```

2. To share the add-in with other people on Mac, send them https://yuxuancao97.github.io/webembed/manifest.xml to copy into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`. Or deploy it org-wide through the Microsoft 365 admin center. Anyone opening the deck without the add-in sees the snapshot image.
3. A site you embed has to allow `https://yuxuancao97.github.io` in its CSP `frame-ancestors` (or allow any site).

Any other HTTPS static host works the same way: upload `public/` and run the sideload with that `ADDIN_URL`. To go back to the local server, run `npm run sideload` without it.

On static hosting the "is this site embeddable?" check is off, because it needs the Node endpoint. A site that refuses embedding then shows as a blank box with no explanation. Everything else works.

## Limits worth knowing

| Thing | Why |
|---|---|
| Some sites show blank or an error (Google, GitHub, most banks and social sites) | They send `X-Frame-Options` or CSP `frame-ancestors`, which forbids embedding. No add-in can override that. The server checks the headers when you press Embed and warns you first. A blocked site shows an **Open in browser** button on the slide instead. |
| `http://` pages are rejected | The add-in runs over HTTPS, and browsers block insecure frames inside it. |
| Inline HTML is capped at ~1.4 MB | Office stores add-in settings in the document with a size cap (documented as 2 MB). For bigger assets (videos, large images), host the files and embed by URL. |
| Inline HTML can't load other local files | It runs from `srcdoc`, so inline your CSS and JS and use `data:` URIs or absolute https URLs for images and fonts. |
| Inline HTML is sandboxed | `allow-scripts allow-popups` only, with no same-origin access. It can't touch the deck, cookies or the add-in. |
| Embeds need the add-in host, and a connection, while presenting | Even inline HTML is rendered by the add-in page, and that page loads `office.js` from Microsoft's CDN. Copying an asset into the deck removes the dependency on the asset's own site, not on the add-in host. Offline presenting hasn't been tested. |
| The box is capped at 1000×1000 pt on insert | PowerPoint's limit for the initial size. Drag to resize afterwards. Set **Page width** (for example Laptop, 1280 px) to lay a page out at desktop width and scale it down to the box. |

## Commands

| Command | What it does |
|---|---|
| `npm start` | HTTPS server on port 3000 (HTTP if no certs yet) |
| `npm run preview` | Plain HTTP server for previewing `http://localhost:3000` in a browser (no PowerPoint) |
| `npm run certs` | Installs and trusts the localhost dev certificate |
| `npm run sideload` / `npm run unsideload` | Installs or removes the manifest in PowerPoint |
| `npm run validate` | Validates `manifest.xml` with Microsoft's validator |
| `npm run icons` | Regenerates the PNG icons |

## Troubleshooting

- **WebEmbed isn't in the Add-ins list.** Run `npm run sideload` again, then quit PowerPoint with ⌘Q (closing the window isn't enough) and reopen it.
- **The box says the add-in couldn't load.** Make sure `npm start` is running and prints `https://`. Then open `https://localhost:3000/embed.html` in Safari: you should get no certificate warning.
- **Changes don't show up.** PowerPoint caches add-in pages. Quit it, clear `~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/` and `~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`, then reopen.
- **Debugging:** right-click inside the add-in and choose **Inspect Element**. To enable it, run `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`, then restart PowerPoint.

## Files

```
manifest.xml          add-in manifest (content add-in, Presentation host)
server.mjs            static server + /api/check framing check, no dependencies
scripts/sideload.mjs  installs the manifest for PowerPoint on Mac
scripts/make-icons.mjs
public/embed.html     the on-slide app
public/embed.js       storage, URL normalising, rendering, setup UI
public/embed.css
public/samples/       orbit.html, kinetic.html, bars.html
public/assets/        icons
```
