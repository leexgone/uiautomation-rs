use uiautomation_derive::map_as;
use uiautomation_derive::EnumConvert;
use windows::Win32::Foundation::GlobalFree;
use windows::Win32::Foundation::HANDLE;
use windows::Win32::Foundation::HGLOBAL;
use windows::Win32::System::DataExchange::CloseClipboard;
use windows::Win32::System::DataExchange::EmptyClipboard;
use windows::Win32::System::DataExchange::EnumClipboardFormats;
use windows::Win32::System::DataExchange::GetClipboardData;
use windows::Win32::System::DataExchange::IsClipboardFormatAvailable;
use windows::Win32::System::DataExchange::OpenClipboard;
use windows::Win32::System::DataExchange::SetClipboardData;
use windows::Win32::System::Memory::GlobalAlloc;
use windows::Win32::System::Memory::GlobalLock;
use windows::Win32::System::Memory::GlobalSize;
use windows::Win32::System::Memory::GlobalUnlock;
use windows::Win32::System::Memory::GMEM_FIXED;
use windows::Win32::System::Memory::GMEM_ZEROINIT;
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

const HANDLE_FORMATS: [ClipboardFormat; 10] = [ClipboardFormat::BITMAP, ClipboardFormat::DSPBITMAP, ClipboardFormat::ENHMETAFILE,
    ClipboardFormat::DSPENHMETAFILE, ClipboardFormat::DSPMETAFILEPICT, ClipboardFormat::DSPTEXT, ClipboardFormat::HDROP, ClipboardFormat::METAFILEPICT,
    ClipboardFormat::PALETTE, ClipboardFormat::PENDATA];

/// Windows clipboard wrapper.
/// 
/// ```
/// let clipboard = uiautomation::clipboards::Clipboard::open().unwrap();
/// clipboard.set_text("hello").unwrap();
/// let text = clipboard.get_text().unwrap();
/// assert_eq!(text, "hello");
/// ```
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
                let memory = GlobalData::from(data);
                let data = memory.lock();
                let str = PWSTR::from_raw(data.data as _);
                str.to_string()?
            }
        };

        Ok(text)
    }

    /// Sets clipboard data in a string format.
    pub fn set_text(&self, text: &str) -> Result<()> {
        let data: Vec<u16> = text.encode_utf16().chain(std::iter::once(0u16)).collect();
        let data = PWSTR::from_raw(data.as_ptr() as *mut u16);
        let data_len = unsafe { data.len() } + 1;   // sizeof `u16`
        let mut memory = GlobalData::alloc(data_len * 2)?;
        let lock_data = memory.lock();
        unsafe {
            std::ptr::copy(data.as_ptr(), lock_data.data as _, data_len);

            EmptyClipboard()?;
            SetClipboardData(ClipboardFormat::UNICODETEXT as _, memory.handle())?;
        }
        memory.forget();    // move ownership into clipboard if succeed.

        Ok(())
    }

    /// Creates a snapshot of the current clipboard contents.
    /// The snapshot contains all data formats available in the clipboard.
    pub fn snapshot(&self) -> Result<Snapshot> {
        let mut datas = Vec::new();

        let mut format = 0u32;
        let mut has_text = false;
        let mut has_image = false;
        loop {
            format = unsafe { EnumClipboardFormats(format) };

            if format == 0 {
                break;
            } else if HANDLE_FORMATS.iter().find(|t| format == **t as _).is_some() {  // ignore bitmap and other format which is stored by handle
                continue;
            } else if format == ClipboardFormat::DIB as _ || format == ClipboardFormat::DIBV5 as _ {    // image data
                if has_image {
                    continue;
                } else {
                    has_image = true;
                }
            } else if format == ClipboardFormat::TEXT as _ || format == ClipboardFormat::UNICODETEXT as _ || format == ClipboardFormat::OEMTEXT as _ {    // text data
                if has_text {
                    continue;
                } else {
                    has_text = true;
                }
            }

            let data = unsafe { GetClipboardData(format)? };
            let data = GlobalData::from(data);
            let data = data.clone();

            datas.push((format, data));
        }

        Ok(Snapshot { datas })
    }

    /// Restores the clipboard contents from a snapshot.
    /// The snapshot should be created using the `snapshot()` method.
    pub fn restore(&self, snapshot: Snapshot) -> Result<()> {
        unsafe { EmptyClipboard()? };
        for (format, mut data) in snapshot.datas.into_iter() {
            unsafe { SetClipboardData(format, data.handle())? };
            data.forget();
        }

        Ok(())
    }
}

impl Drop for Clipboard {
    fn drop(&mut self) {
        let _ = unsafe { CloseClipboard() };
    }
}

#[derive(Debug, Default)]
pub(crate) struct GlobalData {
    memory: HGLOBAL,
    owned: bool
}

impl GlobalData {
    pub fn alloc(size: usize) -> Result<Self> {
        let data = unsafe {
            GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, size)?
        };
        Ok(Self { memory: data, owned: true })
    }

    // pub fn get(&self) -> HGLOBAL {
    //     self.memory
    // }

    pub fn handle(&self) -> Option<HANDLE> {
        if self.memory.is_invalid() {
            None
        } else {
            Some(unsafe { std::mem::transmute(self.memory) })
        }
    }

    pub fn size(&self) -> usize {
        unsafe {
            GlobalSize(self.memory)
        }
    }

    pub fn lock(&self) -> GlobalGuard {
        let data = unsafe { GlobalLock(self.memory)};
        GlobalGuard{
            memory: self.memory,
            data
        }
    }

    pub fn forget(&mut self) {
        self.owned = false;
    }
}

impl From<HGLOBAL> for GlobalData {
    fn from(data: HGLOBAL) -> Self {
        Self { memory: data, owned: false }
    }
}

impl Into<HGLOBAL> for GlobalData {
    fn into(self) -> HGLOBAL {
        let mut gd = self;
        gd.owned = false;
        gd.memory
    }
}

impl From<HANDLE> for GlobalData {
    fn from(data: HANDLE) -> Self {
        let data: HGLOBAL = unsafe { std::mem::transmute(data) };
        data.into()
    }
}

impl AsRef<HGLOBAL> for GlobalData {
    fn as_ref(&self) -> &HGLOBAL {
        &self.memory
    }
}

impl AsMut<HGLOBAL> for GlobalData {
    fn as_mut(&mut self) -> &mut HGLOBAL {
        &mut self.memory
    }
}

impl Clone for GlobalData {
    fn clone(&self) -> Self {
        if self.memory.is_invalid() {
            Self {
                memory: self.memory,
                owned: false
            }
        } else {
            let len = self.size();
            let data = GlobalData::alloc(len).unwrap();
            let src = self.lock();
            let dst = data.lock();
            unsafe { std::ptr::copy(src.data as _, dst.data as _, len) };

            data
        }
    }
}

impl Drop for GlobalData {
    fn drop(&mut self) {
        if self.owned && !self.memory.is_invalid() { 
            unsafe { GlobalFree(Some(self.memory)).unwrap() };
            self.owned = false
        }
    }
}

pub(crate) struct GlobalGuard{
    memory: HGLOBAL,
    pub data: *mut core::ffi::c_void
}

impl Drop for GlobalGuard {
    fn drop(&mut self) {
        unsafe { GlobalUnlock(self.memory).unwrap() }
    }
}

#[derive(Debug, Default)]
pub struct Snapshot {
    datas: Vec<(u32, GlobalData)>
}