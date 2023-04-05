use std::iter::once;

use windows::Win32::UI::Input::KeyboardAndMouse::GetActiveWindow;
use windows::Win32::UI::WindowsAndMessaging::IDNO;
use windows::Win32::UI::WindowsAndMessaging::IDYES;
use windows::Win32::UI::WindowsAndMessaging::MB_ICONERROR;
use windows::Win32::UI::WindowsAndMessaging::MB_ICONINFORMATION;
use windows::Win32::UI::WindowsAndMessaging::MB_ICONQUESTION;
use windows::Win32::UI::WindowsAndMessaging::MB_ICONWARNING;
use windows::Win32::UI::WindowsAndMessaging::MB_OK;
use windows::Win32::UI::WindowsAndMessaging::MB_YESNO;
use windows::Win32::UI::WindowsAndMessaging::MB_YESNOCANCEL;
use windows::Win32::UI::WindowsAndMessaging::MESSAGEBOX_RESULT;
use windows::Win32::UI::WindowsAndMessaging::MESSAGEBOX_STYLE;
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
    message_box(text, caption, MB_OK);
}

/// Shows an infomation message dialog with an `OK` button and a `INFO` icon.
/// 
/// ``` no_run
/// uiautomation::dialogs::show_info("This is my info.", "INFO");
/// ```
pub fn show_info(text: &str, caption: &str) {
    message_box(text, caption, MB_OK | MB_ICONINFORMATION);
}

/// Shows a warning message dialog with an `OK` button and a `WARN` icon.
/// 
/// ``` no_run
/// uiautomation::dialogs::show_warn("This is my warn.", "WARN");
/// ```
pub fn show_warn(text: &str, caption: &str) {
    message_box(text, caption, MB_OK | MB_ICONWARNING);
}

/// Shows an error message dialog with an `OK` button and a `ERROR` icon.
/// 
/// ``` no_run
/// uiautomation::dialogs::show_error("This is my error.", "ERROR");
/// ```
pub fn show_error(text: &str, caption: &str) {
    message_box(text, caption, MB_OK | MB_ICONERROR);
}

/// Shows a query dialog with `Yes` and `No` button. Returns `true` only when user pressed `Yes` button.
/// 
/// ``` no_run
/// use uiautomation::dialogs::*;
/// 
/// if query_yes_no("Are your sure?", "QUERY") {
///     show_message("Oh! Yes!!!", "ANSWERED");
/// }
/// ```
pub fn query_yes_no(text: &str, caption: &str) -> bool {
    message_box(text, caption, MB_YESNO | MB_ICONQUESTION) == IDYES
}

/// Shows a query dialog with `Yes` `No` and `Cancel` button.
/// 
/// Returns `Some(true)` when user pressed `Yes` button, returns `Some(false)` when user pressed `No` button, returns `None` when user pressed `Cancel` button.
/// 
/// ``` no_run
/// use uiautomation::dialogs::*;
/// 
/// match query_yes_no_cancel("Do you want to do this?", "QUERY") {
///     Some(ret) => show_message(if ret { "YES!" } else { "OK!!" }, "ANSWER"),
///     None => show_message("I don't known.", "ANSWER")
/// }
/// ```
pub fn query_yes_no_cancel(text: &str, caption: &str) -> Option<bool> {
    match message_box(text, caption, MB_YESNOCANCEL | MB_ICONQUESTION) {
        IDYES => Some(true),
        IDNO => Some(false),
        _ => None
    }
}

/// Shows a message dialog in front of the active window.
pub fn message_box(text: &str, caption: &str, styles: MESSAGEBOX_STYLE) -> MESSAGEBOX_RESULT {
    let text: Vec<u16> = text.encode_utf16().chain(once(0u16)).collect();
    let lptext = PCWSTR::from_raw(text.as_ptr());
    
    let caption: Vec<u16> = caption.encode_utf16().chain(once(0u16)).collect();
    let lpcaption = PCWSTR::from_raw(caption.as_ptr());

    unsafe {
        let hwnd = GetActiveWindow();
        MessageBoxW(hwnd, lptext, lpcaption, styles)
    }    
}