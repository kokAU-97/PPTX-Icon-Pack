# My Icon Library — a PowerPoint add-in

A task pane add-in that works like PowerPoint's built-in icon gallery (Insert →
Icons), but backed by **your own** icons instead of Microsoft's set. Upload
SVG/PNG/JPG icons once, then click any icon to drop it onto the current slide,
in any presentation, from then on.

## What's in this folder

| File | Purpose |
|---|---|
| `manifest.xml` | Tells PowerPoint the add-in exists and where to load it from. |
| `taskpane.html` / `.css` / `.js` | The task pane UI: upload, search, grid, insert, delete. |
| `commands.html` | Empty support file the manifest schema requires. |
| `assets/icon-16.png`, `icon-32.png`, `icon-80.png` | The add-in's own icon, shown in PowerPoint's ribbon/UI (not your icon library — feel free to replace these with your own branding). |

Icons you add are stored in **IndexedDB inside the task pane**, so they persist
between sessions on that computer/browser profile without any backend server.

## 1. Try it before installing anything

You can open `taskpane.html` directly in a normal browser tab to try the
upload/search/delete/grid UI immediately — everything works except the
"insert into slide" step, which needs to run inside PowerPoint via Office.js.

## 2. Host the files somewhere HTTPS

PowerPoint requires add-ins to be served over HTTPS. The easiest free option
is **GitHub Pages**:

1. Create a new GitHub repo and push this whole folder to it.
2. In the repo's Settings → Pages, set the source to your main branch.
3. GitHub gives you a URL like `https://yourname.github.io/icon-library-addin/`.

(Any static HTTPS host works too — Netlify, Vercel, Azure Static Web Apps,
your own server, etc.)

## 3. Point the manifest at your host

Open `manifest.xml` and replace every occurrence of
`https://YOUR_HOST_HERE` with your real base URL, e.g.:

```
https://yourname.github.io/icon-library-addin
```

That covers the icon URLs, `AppDomains`, `SourceLocation`, and the
`Commands.Url` / `Taskpane.Url` resources.

## 4. Sideload it into PowerPoint

**PowerPoint on the web:**
Insert tab → Add-ins → Upload My Add-in → choose your `manifest.xml`.

**PowerPoint desktop (Windows or Mac):**
Insert tab → My Add-ins → Upload My Add-in (small link near the top of the
dialog) → browse to `manifest.xml`.

Either way, a **My Icons** button appears on the Home tab, which opens the
task pane.

> Sideloaded add-ins like this are visible only to you. To share it with a
> team, an admin can publish it via the Microsoft 365 admin center instead —
> ask if you'd like help with that version.

## 5. Use it

- **+ Add icons** — pick one or more SVG/PNG/JPG files. They're saved
  permanently into the library.
- **Search** — filter by the icon's name (taken from the file name).
- **Click an icon** — inserts it onto the current slide.
- **× on hover** — deletes that one icon.
- **Clear all** — wipes the whole library (asks for confirmation first).

## Notes & limitations

- The library lives in that browser profile's IndexedDB. It won't
  automatically follow you to another computer, and clearing browser data
  will clear it too. If you want the library shared across your devices or
  team, that needs a small backend (say the word if you want that version).
- SVGs are rasterized to a 512×512 PNG the first time you insert them (then
  cached), since that's the most reliable way to get pixel-perfect results
  across PowerPoint versions.
- Want icon categories/tags, drag-to-resize before inserting, or a "recently
  used" row like the native gallery? All straightforward to add on top of
  this — just ask.
