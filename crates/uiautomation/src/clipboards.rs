use uiautomation_derive::map_as;
use uiautomation_derive::EnumConvert;
use windows::Win32::Foundation::HGLOBAL;
use windows::Win32::System::DataExchange::CloseClipboard;
use windows::Win32::System::DataExchange::EmptyClipboard;
use windows::Win32::System::DataExchange::GetClipboardData;
use windows::Win32::System::DataExchange::IsClipboardFormatAvailable;
use windows::Win32::System::DataExchange::OpenClipboard;
use windows::Win32::System::Memory::GlobalLock;
use windows::Win32::System::Memory::GlobalUnlock;
use windows::Win32::UI::Input::KeyboardAndMouse::GetActiveWindow;
use windows_core::PWSTR;

use crate::types::Handle;
use crate::Result;

#[repr(u16)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::System::Ole::CLIPBOARD_FORMAT)]
pub enum ClipboardFormat {
    /// A handle to a bitmap (HBITMAP).
    BITMAP = 2u16,
    /// A memory object containing a BITMAPINFO structure followed by the bitmap bits.
    DIB = 8u16,
    /// A memory object containing a BITMAPV5HEADER structure followed by the bitmap color space information and the bitmap bits.
    DIBV5 = 17u16,
    /// Software Arts' Data Interchange Format.
    DIF = 5u16,
    /// Bitmap display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in bitmap format in lieu of the privately formatted data.
    DSPBITMAP = 130u16,
    /// Enhanced metafile display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in enhanced metafile format in lieu of the privately formatted data.
    DSPENHMETAFILE = 142u16,
    /// Metafile-picture display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in metafile-picture format in lieu of the privately formatted data.
    DSPMETAFILEPICT = 131u16,
    /// Text display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in text format in lieu of the privately formatted data.
    DSPTEXT = 129u16,
    /// A handle to an enhanced metafile (HENHMETAFILE).
    ENHMETAFILE = 14u16,
    /// Start of a range of integer values for application-defined GDI object clipboard formats. The end of the range is CF_GDIOBJLAST.
    GDIOBJFIRST = 768u16,
    /// See CF_GDIOBJFIRST.
    GDIOBJLAST = 1023u16,
    /// A handle to type HDROP that identifies a list of files. An application can retrieve information about the files by passing the handle to the DragQueryFile function.
    HDROP = 15u16,
    /// The data is a handle (HGLOBAL) to the locale identifier (LCID) associated with text in the clipboard. When you close the clipboard, if it contains CF_TEXT data but no CF_LOCALE data, the system automatically sets the CF_LOCALE format to the current input language. You can use the CF_LOCALE format to associate a different locale with the clipboard text.
    LOCALE = 16u16,
    /// Handle to a metafile picture format as defined by the METAFILEPICT structure. When passing a CF_METAFILEPICT handle by means of DDE, the application responsible for deleting hMem should also free the metafile referred to by the CF_METAFILEPICT handle.
    METAFILEPICT = 3u16,
    /// Text format containing characters in the OEM character set. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
    OEMTEXT = 7u16,
    /// Owner-display format. The clipboard owner must display and update the clipboard viewer window, and receive the WM_ASKCBFORMATNAME, WM_HSCROLLCLIPBOARD, WM_PAINTCLIPBOARD, WM_SIZECLIPBOARD, and WM_VSCROLLCLIPBOARD messages. The hMem parameter must be NULL.
    OWNERDISPLAY = 128u16,
    /// Handle to a color palette. Whenever an application places data in the clipboard that depends on or assumes a color palette, it should place the palette on the clipboard as well.
    PALETTE = 9u16,
    /// Data for the pen extensions to the Microsoft Windows for Pen Computing.
    PENDATA = 10u16,
    /// Start of a range of integer values for private clipboard formats. The range ends with CF_PRIVATELAST. Handles associated with private clipboard formats are not freed automatically; the clipboard owner must free such handles, typically in response to the WM_DESTROYCLIPBOARD message.
    PRIVATEFIRST = 512u16,
    /// See CF_PRIVATEFIRST.
    PRIVATELAST = 767u16,
    /// Represents audio data more complex than can be represented in a CF_WAVE standard wave format.
    RIFF = 11u16,
    /// Microsoft Symbolic Link (SYLK) format.
    SYLK = 4u16,
    /// Text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. Use this format for ANSI text.
    TEXT = 1u16,
    /// Tagged-image file format.
    TIFF = 6u16,
    /// Unicode text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
    UNICODETEXT = 13u16,
    /// Represents audio data in one of the standard wave formats, such as 11 kHz or 22 kHz PCM.
    WAVE = 12u16,
}

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

    /// Empties the clipboard and frees handles to data in the clipboard.
    /// This method will clear the clipboard contents.
    pub fn empty(&self) -> Result<()> {
        unsafe { EmptyClipboard()? };
        Ok(())
    }

    /// Determines whether the clipboard contains data in the specified format.
    pub fn is_format_available(&self, format: ClipboardFormat) -> Result<bool> {
        let result = unsafe {
            IsClipboardFormatAvailable(format as _)   
        };

        match result {
            Ok(_) => Ok(true),
            Err(e) => {
                if e.code().is_ok() {
                    Ok(false)
                } else {
                    Err(e.into())
                }
            }
        }
    }

    /// Retrieves data from the clipboard in a string format. 
    pub fn get_text(&self) -> Result<String> {
        let text = unsafe { 
            IsClipboardFormatAvailable(ClipboardFormat::UNICODETEXT as _)?;

            let data = GetClipboardData(ClipboardFormat::UNICODETEXT as _)?;
            if data.is_invalid() {
                String::default()
            } else {
                let memory: HGLOBAL = std::mem::transmute(data);
                let buffer = GlobalLock(memory);
                let str = PWSTR::from_raw(buffer as _);
                let ret = str.to_string();
                GlobalUnlock(memory)?;

                ret?
            }
        };

        Ok(text)
    }
}

impl Drop for Clipboard {
    fn drop(&mut self) {
        let _ = unsafe { CloseClipboard() };
    }
}