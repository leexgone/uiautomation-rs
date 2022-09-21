use std::convert::TryFrom;

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

    fn try_from(handle: Handle) -> Result<Self> {
        let hdc = unsafe {
            GetDC(handle)    
        };

        if hdc.is_invalid() {
            Ok(Canvas {
                window: Some(handle),
                hdc: hdc
            })
        } else {
            Err(Error::new(ERR_INVALID_OBJECT, "Invalid dc object"))
        }
    }
}

impl Canvas {
    pub fn is_valid(&self) -> bool {
        self.hdc.is_invalid()
    }
}

// impl Drop for Canvas {
//     fn drop(&mut self) {
//         if self.is_valid() {
//             match self.window {
//                 Some(hwnd) => unsafe { ReleaseDC(hwnd, self.hdc); },
//                 None => unsafe { DeleteDC(self.hdc); }
//             }
            
//         }
//     }
// }