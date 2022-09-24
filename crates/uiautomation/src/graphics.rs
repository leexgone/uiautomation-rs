use std::convert::TryFrom;

use windows::Win32::Graphics::Gdi::CreateCompatibleDC;
use windows::Win32::Graphics::Gdi::CreatedHDC;
use windows::Win32::Graphics::Gdi::DeleteDC;
use windows::Win32::Graphics::Gdi::GetDC;
use windows::Win32::Graphics::Gdi::HDC;
use windows::Win32::Graphics::Gdi::ReleaseDC;

use crate::Error;
use crate::Result;
use crate::errors::ERR_INVALID_OBJECT;
use crate::types::Handle;

/// Canvas is a wrapper for `HDC` object which can used as a gdi canvas.
#[derive(Debug)]
pub struct Canvas{
    window: Option<Handle>,
    hdc: HDC
}

impl TryFrom<Handle> for Canvas {
    type Error = Error;

    fn try_from(wnd_handle: Handle) -> Result<Self> {
        let hdc = unsafe {
            GetDC(wnd_handle)    
        };

        if hdc.is_invalid() {
            Ok(Canvas {
                window: Some(wnd_handle),
                hdc: hdc
            })
        } else {
            Err(Error::new(ERR_INVALID_OBJECT, "Invalid dc object"))
        }
    }
}

impl Canvas {
    // pub fn create() -> Result<Canvas> {
    //     CreateCompatibleDC(hdc)
    // }

    pub fn is_valid(&self) -> bool {
        self.hdc.is_invalid()
    }
}

impl Drop for Canvas {
    fn drop(&mut self) {
        if self.is_valid() {
            match self.window {
                Some(hwnd) => unsafe { ReleaseDC(hwnd, self.hdc); },
                None => {
                    let hdc = CreatedHDC(self.hdc.0);
                    unsafe { DeleteDC(hdc) }; 
                }
            }
        }
    }
}