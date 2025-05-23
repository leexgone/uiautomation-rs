use windows::Win32::System::DataExchange::CloseClipboard;
use windows::Win32::System::DataExchange::EmptyClipboard;
use windows::Win32::System::DataExchange::OpenClipboard;
use windows::Win32::UI::Input::KeyboardAndMouse::GetActiveWindow;

use crate::types::Handle;
use crate::Result;

/// Windows clipboard wrapper.
#[derive(Debug)]
pub struct Clipboard {
}

impl Clipboard {
    const MAX_OPEN_TIMES: u32 = 10;

    fn create(owner: Option<Handle>) -> Result<Self> {
        unsafe { OpenClipboard(owner.map(|h| h.into()))? };
        Ok(Self {})
    }

    fn attempt(owner: Option<Handle>) -> Result<Self> {
        let mut times = Self::MAX_OPEN_TIMES;

        loop {
            let result = Self::create(owner);
            if result.is_ok() || times <= 0 {
                break result;
            } else {
                std::thread::sleep(std::time::Duration::from_millis(5));
                times -= 1;
            }
        }
    }

    /// Create a new clipboard instance.
    pub fn new() -> Result<Self> {
        Self::create(None)
    }

    /// Create a new clipboard instance with the specified owner window.
    pub fn new_for(owner: Handle) -> Result<Self> {
        Self::create(Some(owner))
    }

    /// Open the clipboard for the current active window.
    /// This method will try to open the clipboard multiple times if it fails.
    pub fn open() -> Result<Self> {
        let window = {
            let wnd = unsafe { GetActiveWindow() };
            if wnd.is_invalid() {
                None
            } else {
                Some(Handle::from(wnd))
            }
        };
        Self::attempt(window)
    }

    /// Open the clipboard for the specified owner window.
    /// This method will try to open the clipboard multiple times if it fails.
    pub fn open_for(owner: Handle) -> Result<Self> {
        Self::attempt(Some(owner))
    }

    /// Empty the clipboard.
    /// This method will clear the clipboard contents.
    pub fn empty() -> Result<()> {
        unsafe { EmptyClipboard()? };
        Ok(())
    }

    // pub fn get_text(&self) -> Result<String> {
    //     CF_TEXT
    // }
}

impl Drop for Clipboard {
    fn drop(&mut self) {
        let _ = unsafe { CloseClipboard() };
    }
}