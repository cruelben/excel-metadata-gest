# GIF Frame Master

> A powerful browser-based GIF editor with frame-by-frame control, visual crop, frame reordering, GIF appending, duplicate detection, and custom export.

![Status](https://img.shields.io/badge/status-active-brightgreen)
![License](https://img.shields.io/badge/license-MIT-blue)
![Made with](https://img.shields.io/badge/made%20with-vanilla%20JS-yellow)

---

## 📖 Overview

**GIF Frame Master** is a lightweight, dependency-free (except for two decoding/encoding libraries) web app and Chrome extension for editing animated GIFs directly in the browser.

Unlike most online GIF editors, everything runs **locally** — no uploads, no servers, no tracking. Just open the page and start editing.

### Key features

- 🎞️ **Frame-by-frame management** — select, reorder, remove
- 🖱️ **Drag & drop reordering** — move single frames or whole selections
- 🔍 **Duplicate detection** — find near-identical consecutive frames with adjustable threshold
- ✂️ **Visual crop** — drag a rectangle to crop every frame at once
- ➕ **Append multiple GIFs** — queue a second GIF, position and scale it visually, then merge
- 🎬 **Animated preview** — real-time preview with original delays before exporting
- ⚡ **Speed control** — slow down or speed up the whole animation
- 🔄 **Ping-pong effect** — forward-then-backward playback
- 🎚️ **Quality / size slider** — trade off between file size and quality
- ↩️ **Undo** — up to 20 actions reversible
- 🌍 **English UI** — clean, minimal, dark theme

---

## 🚀 Installation

### Option 1 — Use it as a web app

1. Download or clone this repository:
   ```bash
   git clone https://github.com/cruelben/GIF-FrameMaster.git
   ```
2. Open `index.html` in any modern browser (Chrome, Edge, Firefox, Brave).
3. That's it. No build step, no server required.

### Option 2 — Install as a Chrome extension

1. Clone or download this repository.
2. Open Chrome and go to `chrome://extensions/`.
3. Enable **Developer mode** (top-right toggle).
4. Click **Load unpacked** and select the folder of this repository.
5. The GIF Frame Master icon appears in your toolbar. Click it to open the editor in a new tab.

> Works identically on any Chromium-based browser (Edge, Brave, Opera, Vivaldi).

---

## 🎮 How to use

### 1. Load a GIF
Click **Select GIF** and choose an animated GIF. The app decodes every frame and shows them in a grid.

### 2. Manage frames
- **Click** a frame to select it (red border)
- **Shift + click** to select a range
- **Drag & drop** a selection to reorder
- **Double-click** a frame to view it at real size in a new tab
- Use the toolbar to: select all, deselect all, invert selection, undo, decimate (remove 1 every N), detect duplicates, remove selected from view

### 3. Append a second GIF *(optional)*
Click **➕ Add GIF to queue** in the Load tab, pick a second GIF, then:
- Drag to position it inside the fixed canvas
- Use **Zoom** to scale it
- Choose a **background color** for uncovered areas
- Use **Fit: cover** or **Fit: contain** presets
- Confirm to append the frames at the end of the current sequence

### 4. Crop
In the **Crop** tab, drag the rectangle to select the area you want to keep. Every frame will be cropped accordingly.

### 5. Export
In the **Export** tab:
- Adjust **speed** (0.5x – 2x)
- Toggle **ping-pong** effect
- Adjust **quality/size** slider (1 = best, 20 = smallest)
- Click **Create new GIF** and download the result

---

## 🧱 Project structure

```
GIF-FrameMaster/
├── index.html          # UI: layout, tabs, styles
├── app.js              # All application logic
├── background.js       # Chrome extension service worker
├── manifest.json       # Chrome extension manifest (MV3)
├── icon.png            # Extension / header icon
├── gifuct-js.min.js    # GIF decoder (parsing frames)
├── gifshot.min.js      # GIF encoder (export)
└── README.md
```

### Architecture at a glance

- **Decoding** — [`gifuct-js`](https://github.com/matt-way/gifuct-js) parses the GIF into raw frames + disposal metadata.
- **Composition** — `composeAllFrames()` reconstructs each full frame by honoring disposal types 2 (clear) and 3 (restore), producing independent canvases.
- **Editing** — selection, reorder, dedupe, crop, and append all operate on this array of `{ canvas, delay, disposalType }`.
- **Preview** — a lightweight canvas player cycles through active frames with their real delays.
- **Encoding** — [`gifshot`](https://github.com/yahoo/gifshot) rebuilds the final animated GIF from cropped canvas frames.

No frameworks. No bundler. Just vanilla JavaScript and the DOM.

---

## 🛠️ Tech stack

| Layer | Technology |
|-------|-----------|
| UI | HTML5 + CSS3 (custom, no framework) |
| Logic | Vanilla JavaScript (ES2017+) |
| GIF decoding | gifuct-js |
| GIF encoding | gifshot |
| Storage (extension) | chrome.storage (reserved for future use) |
| Manifest | Chrome Extension Manifest V3 |

---

## 🗺️ Roadmap

- [x] Frame grid with multi-selection
- [x] Drag & drop reordering
- [x] Duplicate detection with adjustable threshold
- [x] Visual crop
- [x] Append GIF with position / scale / background
- [x] Animated preview
- [x] Speed control & ping-pong
- [x] Robust error handling with technical details
- [ ] Session save & restore (`chrome.storage`)
- [ ] Export individual frames as PNG
- [ ] Per-frame delay editing
- [ ] Drag & drop GIF file onto the page
- [ ] Light theme
- [ ] Keyboard shortcuts

---

## 🤝 Contributing

Contributions are welcome. To contribute:

1. Fork this repository
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Commit your changes (`git commit -m "Add my feature"`)
4. Push to your branch (`git push origin feature/my-feature`)
5. Open a Pull Request

### Coding conventions

- **UI strings:** English
- **Code comments:** Italian (author's preferred language)
- **Style:** vanilla JS, no frameworks, no bundler
- **Indent:** 4 spaces
- **Naming:** camelCase for variables/functions, UPPER_SNAKE_CASE for constants

Please test your changes manually (open `index.html` in a browser or reload the unpacked extension) before submitting.

---

## 🐛 Known limitations

- Very large GIFs (hundreds of frames at high resolution) may be slow or hit browser memory limits during export.
- The encoder (gifshot) runs in the main thread; future versions may move to Web Workers.
- Disposal method 3 (restore to previous) is supported but rarely used by encoders.

---

## 📄 License

This project is released under the **MIT License**. You are free to use, modify, and distribute it, including for commercial purposes.

> The bundled libraries `gifuct-js` and `gifshot` retain their own licenses. Please refer to their respective repositories.

---

## 👤 Author

**Bruno "cruelben"**
- 🌐 Website: [cruelben.github.io](https://cruelben.github.io/)
- 💻 GitHub: [@cruelben](https://github.com/cruelben)

> *"Surviving Cobol programmer (yes, we still exist!), cryptography systems, and tech enthusiast."*

---

## ⭐ Show your support

If GIF Frame Master saved you time, consider:
- ⭐ Starring this repository
- 🐛 Reporting bugs or suggesting features via [Issues](https://github.com/cruelben/GIF-FrameMaster/issues)
- 🔀 Submitting a Pull Request

---

<p align="center">
  <sub>Built with care, without frameworks.</sub>
</p>