# @cw/desktop — WordCanvas as a Tauri desktop app

Wraps the `@forevka/wordcanvas` editor in a [Tauri v2](https://tauri.app) webview
shell. The editor runs **fully offline** (no `backendUrl` → no network, sync, or
share), and file open/export go through **native OS dialogs**. Ships as a single
**portable `.exe`** (no installer).

## How it's wired

- The frontend (`index.html` + `src/`) is a tiny Vite app that consumes the
  editor as a **pre-built, external** bundle — same pattern as
  `examples/embed-live` (import map + `rollupOptions.external` + `copy-lib`). The
  editor's own import/export workers and ~9 MB of fonts ship inside that bundle.
- The Rust shell (`src-tauri/`) registers `tauri-plugin-dialog` and two small
  commands — `read_file_bytes` / `write_file_bytes` — that the JS bridge
  (`src/native.ts`) uses to read/write the path the dialog returns. This avoids
  fs-plugin scope configuration entirely.
- Open: **Ctrl/Cmd+O** or **File → Open…** → `editor.openDocx(bytes)`.
- Save: the editor's built-in **Export DOCX/PDF** buttons (and **File → Save
  as…**) route through the `onSave` hook → native Save dialog.

## Prerequisites

- **Node 22** and a working repo `npm install` at the root.
- **Rust toolchain** (`rustup`, stable) — required to compile the Tauri shell.
- **WebView2** runtime — preinstalled on Windows 10/11. A portable exe assumes
  it's present (no installer = no WebView2 bootstrapper fallback).

## Develop

From the **repo root** (one-time, so the external editor bundle exists):

```sh
npm install
npm run build:lib --workspace @forevka/wordcanvas
```

Then run the desktop app in dev (hot-reloads the frontend; Rust rebuilds on change):

```sh
npm run desktop:dev          # from repo root
# or:  cd desktop && npm run tauri:dev
```

`npm run dev` alone (without Tauri) opens the frontend in a browser — the native
menu/dialogs are inert there, the editor's default browser download applies.

## Build the portable .exe

```sh
npm run build:lib --workspace @forevka/wordcanvas   # if not already built
npm run desktop:build                               # from repo root
```

The binary lands at `desktop/src-tauri/target/release/wordcanvas-desktop.exe` —
a self-contained portable executable. CI (`.github/workflows/desktop-release.yml`)
builds the same thing on `windows-latest` and attaches it to a draft release on a
`desktop-v*` tag.

## Notes / follow-ups

- **No icon configured** — Tauri uses its default. To customize:
  `cd desktop && npx tauri icon path/to/icon.png` (generates `src-tauri/icons/`,
  then add an `"icon"` array under `bundle` in `tauri.conf.json`).
- **CSP is disabled** (`app.security.csp: null`) for the scaffold. Tighten before
  shipping if desired.
- File bytes cross the IPC boundary as a number array — fine for typical
  documents; switch to `tauri-plugin-fs` or a raw IPC response if you handle very
  large files.
- **Cross-platform**: `bundle.active: false` + `windows_subsystem`/`crt-static`
  are Windows-centric. For macOS/Linux you'd add an OS matrix and likely want
  `bundle.active: true` with `.app`/`.dmg`/`.AppImage` targets instead of a raw
  binary.
