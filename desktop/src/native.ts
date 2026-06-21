// Native save bridge for the Tauri shell.
//
// The picker uses tauri-plugin-dialog (JS). The bytes are written by a tiny Rust
// command (`write_file_bytes`, see src-tauri/src/lib.rs) that uses std::fs on the
// path the dialog returns. This deliberately avoids the tauri-plugin-fs scope
// dance: the app's own command isn't gated by capability scopes, so any
// user-chosen path just works.
//
// Open isn't here on purpose: the editor's ribbon Open uses <input type=file>,
// which is the native OS picker in the webview and already shows the editor's
// import progress overlay — no custom open path needed.

import { save } from "@tauri-apps/plugin-dialog";
import { invoke } from "@tauri-apps/api/core";

export type SaveFormat = "docx" | "pdf";

/** True when running inside the Tauri webview (vs a plain browser dev server). */
export function isTauri(): boolean {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
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
