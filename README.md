<div align="center">

# 📝 Form Studio

### Design a template. Fill it. Export a finished form.

**HTML · CSS · JavaScript · Canvas-style editing · PWA**

[Features](#-features) · [Quick start](#-quick-start) · [Field types](#-supported-form-objects) · [Shortcuts](#-keyboard-shortcuts) · [Deployment](#-deployment)

</div>

Form Studio is a browser-based form template builder with a linked filler page. Place and arrange fields on a paper-sized canvas, save the layout as JSON, then enter values and export the completed form as PDF, PNG, or JPG.

## ✨ Features

| Builder | Filler |
|---|---|
| Drag-and-drop fields with alignment guides and snap behavior | Fill fields on the same sheet layout |
| Resize objects using corner and edge handles | Text, selections, dates, files, signatures, and ratings |
| Edit labels, required flags, position, and dimensions | Native browser validation for applicable inputs |
| Layers, ordering, and lock/unlock controls | Load the latest local template or upload JSON |
| Undo/redo, duplicate, clear sheet, and keyboard shortcuts | Export PDF, PNG, and JPG |
| Auto-save and restore the browser draft | Standard, High, or Print (300 DPI) export quality |
| Collapsible, resizable palette and properties panels | Collapsible, resizable template information sidebar |

**Paper sizes:** A5 · A4 · A3 · Letter · Legal · Tabloid · Executive, with portrait and landscape orientation.

## 🚀 Quick start

1. Clone or download this repository.
2. Open [index.html](index.html) in a modern browser to launch the builder.
3. Choose a template name, paper size, and orientation.
4. Drag fields onto the sheet, adjust their properties, and arrange layers.
5. Download the template JSON as a portable copy.
6. Choose **Go To Filler**, or open [filler.html](filler.html).
7. Fill the form, choose export quality, and download PDF, PNG, or JPG.

The project is static: no package installation or backend server is required for the basic pages. Use an HTTP server or HTTPS hosting to exercise service worker and PWA behavior. The desktop/laptop layout provides the most room for editing.

## 🧩 Supported form objects

| Group | Objects |
|---|---|
| Layout & structure | Section Header, Divider Line, Constant/Label |
| Text | Short Text, Email, Phone, URL, Long Text/Paragraph |
| Numbers | Number, Currency, Percentage |
| Selection | Dropdown/Select, Radio Buttons, Checkboxes, Multi-Select |
| Date & time | Date Picker, Time Picker, Date & Time |
| Files & media | File Upload, Photo Upload |
| Special | Signature drawing field, Rating (1–10) |

The filler aligns inputs with the template. Export outputs include constants and filled values; empty fields are omitted from the exported result.

## ⌨️ Keyboard shortcuts

| Action | Windows/Linux | macOS |
|---|---|---|
| Remove selected object | Delete or Backspace | Delete or Backspace |
| Duplicate | Ctrl + D | Cmd + D |
| Undo | Ctrl + Z | Cmd + Z |
| Redo | Ctrl + Y | Cmd + Shift + Z |
| Nudge 1 px | Arrow keys | Arrow keys |
| Nudge 10 px | Shift + Arrow keys | Shift + Arrow keys |

## 💾 Templates and browser storage

- Builder drafts are saved in **localStorage** and restored on reload.
- The filler reads the latest template from the same browser storage, or accepts a manually uploaded JSON template.
- Browser data is tied to its browser profile and origin. Export JSON to move templates between devices or retain a backup before clearing browser data.
- The current project documents a local builder/filler workflow. A hosted submission API and publish/versioning workflow are future work.

## 📁 Repository guide

| File | Purpose |
|---|---|
| [index.html](index.html), [app.js](app.js), [styles.css](styles.css) | Template builder and canvas controls |
| [filler.html](filler.html), [filler.js](filler.js), [filler.css](filler.css) | Form entry and export workflow |
| [manifest.webmanifest](manifest.webmanifest), [pwa.js](pwa.js), [sw.js](sw.js) | Installation and service worker support |
| [icon.svg](icon.svg) | App icon |

## 📱 Progressive web app

The manifest and service worker provide an installable app shell in compatible browsers, standalone display, configured icon/theme color, and caching for key app files. The existing service worker also supports runtime caching for same-origin files and CDN libraries used by exports.

Installation requires a supported browser and HTTPS (or localhost). After deployment, reopen the app and refresh to pick up the latest assets. Export features that depend on CDN libraries may need an initial online visit before those libraries are available offline.

## 🌐 Deployment

Form Studio can be hosted as a static website. For GitHub Pages:

1. Open repository **Settings → Pages**.
2. Choose **Deploy from a branch**, then **main** and **/ (root)**.
3. Save and use the URL provided by GitHub.

The repository's documented Pages path is `https://vignesh-s-github.github.io/Form-Studio/`. Use the actual deployment URL shown in Pages settings; service worker behavior is designed to work under the project path.

## 🔎 Manual review during development

After editing the app, check template creation, save/restore, builder-to-filler navigation, and PDF/PNG/JPG exports at each quality level. Review resizing, layer controls, undo/redo, and keyboard shortcuts. Use a fresh browser reload when cached CSS or JavaScript appears out of date.

## 🛤️ Future improvements

- Reusable sections and header/footer areas
- Rich validation rules: patterns, ranges, and file limits
- Conditional logic between fields
- Template publishing and version history
- Submission API and a richer filler runtime

---

<p align="center"><sub>From a blank sheet to a reusable form workflow.</sub></p>
