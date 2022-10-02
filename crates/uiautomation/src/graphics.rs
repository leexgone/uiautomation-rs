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

#[derive(Debug)]
enum HDCOrigin {
    Created(CreatedHDC),
    GotFrom(Handle)    
}

/// Canvas is a wrapper for `HDC` object which can used as a gdi canvas.
#[derive(Debug)]
pub struct Canvas{
    origin: HDCOrigin,
    hdc: HDC
}

impl TryFrom<Handle> for Canvas {
    type Error = Error;

    fn try_from(wnd_handle: Handle) -> Result<Self> {
        let hdc = unsafe {
            GetDC(wnd_handle)    
        };

        if !hdc.is_invalid() {
            Ok(Canvas {
                origin: HDCOrigin::GotFrom(wnd_handle),
                hdc
            })
        } else {
            Err(Error::new(ERR_INVALID_OBJECT, "Invalid dc object"))
        }
    }
}

impl Canvas {
    /// Determines whether the canvas is valid.
    pub fn is_valid(&self) -> bool {
        !self.hdc.is_invalid()
    }

    /// Creates a memory canvas compatible with the current canvas.
    pub fn create_compatible_canvas(&self) -> Result<Canvas> {
        let hdc  = unsafe {
            CreateCompatibleDC(self.hdc)
        };

        if !hdc.is_invalid() {
            Ok(Canvas { 
                origin: HDCOrigin::Created(hdc), 
                hdc: HDC(hdc.0) 
            })
        } else {
            Err(Error::new(ERR_INVALID_OBJECT, "Invalid dc object"))
        }
    }
}

impl Drop for Canvas {
    fn drop(&mut self) {
        if self.is_valid() {
            match self.origin {
                HDCOrigin::Created(cdc) => unsafe {
                    DeleteDC(cdc);
                },
                HDCOrigin::GotFrom(wnd) => unsafe {
                    ReleaseDC(wnd, self.hdc);
                }
            }
            self.hdc = HDC::default();
        }
    }
}