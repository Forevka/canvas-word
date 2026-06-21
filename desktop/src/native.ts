// Native file I/O bridge for the Tauri shell.
//
// Pickers use tauri-plugin-dialog (JS). The actual bytes are read/written by two
// tiny Rust commands (`read_file_bytes` / `write_file_bytes`, see
// src-tauri/src/lib.rs) which use std::fs on the path the dialog returns. This
// deliberately avoids the tauri-plugin-fs scope dance: the app's own commands
// aren't gated by capability scopes, so any user-chosen path just works.
//
// Everything here no-ops outside Tauri (e.g. plain `vite dev` in a browser), so
// the same entry runs in both — see `isTauri()`.

import { open, save } from "@tauri-apps/plugin-dialog";
import { invoke } from "@tauri-apps/api/core";

export type SaveFormat = "docx" | "pdf";

/** True when running inside the Tauri webview (vs a plain browser dev server). */
export function isTauri(): boolean {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
}

/** Native open dialog → file bytes, or null if the user cancelled. */
export async function openDocxViaDialog(): Promise<ArrayBuffer | null> {
  const selected = await open({
    multiple: false,
    directory: false,
    filters: [{ name: "Word document", extensions: ["docx"] }],
  });
  if (typeof selected !== "string") return null;
  // number[] over IPC; fine for typical .docx sizes. Optimize later if needed.
  const data = await invoke<number[]>("read_file_bytes", { path: selected });
  return new Uint8Array(data).buffer;
}

/** Native save dialog → write bytes to the chosen path. No-op if cancelled. */
export async function saveBytesViaDialog(bytes: Uint8Array, format: SaveFormat): Promise<void> {
  const path = await save({
    defaultPath: `document.${format}`,
    filters: [{ name: format.toUpperCase(), extensions: [format] }],
  });
  if (!path) return;
  await invoke("write_file_bytes", { path, data: Array.from(bytes) });
}
