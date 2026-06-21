use std::fs;

// Write exported bytes to the path the user picked via the native save dialog.
// The app's own command isn't gated by fs-plugin scopes, so any chosen path
// works without scope configuration. (Open needs no command — the ribbon's
// <input type=file> is the native OS picker in the webview.)
#[tauri::command]
fn write_file_bytes(path: String, data: Vec<u8>) -> Result<(), String> {
    fs::write(&path, &data).map_err(|e| e.to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![write_file_bytes])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
