use std::convert::TryFrom;

use windows::Win32::Graphics::Gdi::BitBlt;
use windows::Win32::Graphics::Gdi::CreateCompatibleDC;
use windows::Win32::Graphics::Gdi::CreatedHDC;
use windows::Win32::Graphics::Gdi::DeleteDC;
use windows::Win32::Graphics::Gdi::GetDC;
use windows::Win32::Graphics::Gdi::GetStretchBltMode;
use windows::Win32::Graphics::Gdi::HDC;
use windows::Win32::Graphics::Gdi::ROP_CODE;
use windows::Win32::Graphics::Gdi::ReleaseDC;

use windows::Win32::Graphics::Gdi::STRETCH_BLT_MODE;
use windows::Win32::Graphics::Gdi::SetStretchBltMode;
use windows::Win32::Graphics::Gdi::StretchBlt;

use crate::Error;
use crate::Result;
use crate::errors::ERR_INVALID_OBJECT;
use crate::types::Handle;
use crate::types::Point;
use crate::types::Rect;

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

    /// Sets the bitmap stretching mode in the specified device context.
    pub fn set_stretch_blt_mode(&self, mode: STRETCH_BLT_MODE) -> Result<STRETCH_BLT_MODE> {
        let mode = unsafe {
            SetStretchBltMode(self.hdc, mode)
        };

        if mode <= 0 {
            Err(Error::from("Failed to set stretch blt mode"))
        } else {
            Ok(STRETCH_BLT_MODE(mode as _))
        }
    }

    /// Retrieves the current stretching mode. 
    pub fn get_stretch_blt_mode(&self) -> Result<STRETCH_BLT_MODE> {
        let mode = unsafe {
            GetStretchBltMode(self.hdc)
        };

        if mode <= 0 {
            Err(Error::from("Failed to get stretch blt mode"))
        } else {
            Ok(STRETCH_BLT_MODE(mode as _))
        }
    }

    /// The StretchBlt function copies a bitmap from a source rectangle into a destination rectangle.
    pub fn stretch_blt(&self, dest_rect: Rect, src_canvas: &Canvas, src_rect: Rect, rop: ROP_CODE) -> Result<()> {
        let ret = unsafe {
            StretchBlt(self.hdc, dest_rect.get_left(), dest_rect.get_top(), dest_rect.get_width(), dest_rect.get_height(), 
                src_canvas.hdc, src_rect.get_left(), src_rect.get_top(), src_rect.get_width(), src_rect.get_height(), rop)
        };

        if ret.as_bool() {
            Ok(())
        } else {
            Err(Error::from("Failed to stretch blt"))
        }
    }

    /// Performs a bit-block transfer of the color data corresponding to a rectangle of pixels 
    /// from the specified source device context into a destination device context.
    pub fn bit_blt(&self, dest_rect: Rect, src_canvas: &Canvas, src_pt: Point, rop: ROP_CODE) -> Result<()> {
        let ret = unsafe {
            BitBlt(self.hdc, dest_rect.get_left(), dest_rect.get_top(), dest_rect.get_width(), dest_rect.get_height(), 
                src_canvas.hdc, src_pt.get_x(), src_pt.get_y(), rop)
        };

        if ret.as_bool() {
            Ok(())
        } else {
            Err(Error::from("Failed to bitblt"))
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