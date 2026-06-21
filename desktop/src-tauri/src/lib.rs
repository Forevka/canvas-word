use std::fs;

// Read an arbitrary file the user picked via the native open dialog. The dialog
// (JS side) returns a path; the app's own command isn't gated by fs-plugin
// scopes, so any chosen path works without scope configuration.
#[tauri::command]
fn read_file_bytes(path: String) -> Result<Vec<u8>, String> {
    fs::read(&path).map_err(|e| e.to_string())
}

// Write exported bytes to the path the user picked via the native save dialog.
#[tauri::command]
fn write_file_bytes(path: String, data: Vec<u8>) -> Result<(), String> {
    fs::write(&path, &data).map_err(|e| e.to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![read_file_bytes, write_file_bytes])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
