// Prevents an extra console window on Windows in release builds — important for a
// GUI portable .exe. DO NOT REMOVE.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    wordcanvas_desktop_lib::run()
}
