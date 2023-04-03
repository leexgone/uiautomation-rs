use std::iter::once;

use windows::Win32::UI::Input::KeyboardAndMouse::GetActiveWindow;
use windows::Win32::UI::WindowsAndMessaging::MB_OK;
use windows::Win32::UI::WindowsAndMessaging::MessageBoxW;
use windows::core::PCWSTR;

/// Shows a simple message dialog with only an `OK` button.
/// 
/// ``` no_run
/// use uiautomation::dialogs::show_message;
/// 
/// show_message("Hello, I'm OK.", "INFO");
/// show_message("My Message.", "Test");
/// ```
pub fn show_message(text: &str, caption: &str) {
    let text: Vec<u16> = text.encode_utf16().chain(once(0u16)).collect();
    let lptext = PCWSTR::from_raw(text.as_ptr());
    
    let caption: Vec<u16> = caption.encode_utf16().chain(once(0u16)).collect();
    let lpcaption = PCWSTR::from_raw(caption.as_ptr());

    unsafe {
        let hwnd = GetActiveWindow();
        MessageBoxW(hwnd, lptext, lpcaption, MB_OK);
    }
}