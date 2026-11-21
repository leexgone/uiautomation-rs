use std::path::Path;

use windows::Win32::Graphics::Gdi::{
    BitBlt, CreateCompatibleBitmap, CreateCompatibleDC, CreateDCW,
    DeleteDC, DeleteObject, GetDIBits, GetObjectW, SelectObject,
    BITMAP, BITMAPINFO, BITMAPINFOHEADER, DIB_RGB_COLORS, SRCCOPY,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetSystemMetrics,
    SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN,
    SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN,
};
use windows_core::w;

use crate::types::Rect;
use crate::errors::ERR_INVALID_ARG;
use crate::Error;
use crate::Result;

use super::UIElement;

/// Pixel format of the captured screenshot data.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PixelFormat {
    /// Raw GDI output: 4 bytes per pixel (BGRA), bottom-up (row 0 is the bottom of the image).
    BGRA,
    /// Top-down row order with RGBA channel layout. Suitable for most image libraries.
    RGBA,
}

/// A captured screenshot containing pixel data and metadata.
///
/// The pixel data is stored in BGRA format (bottom-up) by default.
/// Use [`Screenshot::to_rgba()`] to convert to top-down RGBA format.
#[derive(Debug)]
pub struct Screenshot {
    pixels: Vec<u8>,
    width: u32,
    height: u32,
    format: PixelFormat,
}

impl Screenshot {
    /// Creates a new screenshot from raw pixel data.
    pub fn new(pixels: Vec<u8>, width: u32, height: u32) -> Self {
        Self {
            pixels,
            width,
            height,
            format: PixelFormat::BGRA,
        }
    }

    /// Captures the entire desktop across all monitors.
    ///
    /// This method unions all monitor rectangles and captures the combined area.
    /// Negative coordinates are supported for multi-monitor setups.
    pub fn capture_desktop() -> Result<Self> {
        let desktop_rect = get_virtual_screen_rect()?;
        Self::capture_rect(desktop_rect)
    }

    /// Captures a rectangular region of the desktop in screen coordinates.
    ///
    /// The `rect` parameter uses screen coordinates (pixels), where negative
    /// values are valid for multi-monitor setups with secondary displays.
    pub fn capture_rect(rect: Rect) -> Result<Self> {
        let width = rect.get_right() - rect.get_left();
        let height = rect.get_bottom() - rect.get_top();

        if width <= 0 || height <= 0 {
            return Err(Error::new(ERR_INVALID_ARG, "screenshot rect has zero or negative size"));
        }

        let width = width as u32;
        let height = height as u32;

        // Create a DC for the entire virtual screen
        let hdc_screen = unsafe { CreateDCW(w!("DISPLAY"), None, None, None) };
        if hdc_screen.is_invalid() {
            return Err(Error::last_os_error());
        }

        let hdc_mem = unsafe { CreateCompatibleDC(Some(hdc_screen)) };
        if hdc_mem.is_invalid() {
            unsafe { let _ = DeleteDC(hdc_screen); }
            return Err(Error::last_os_error());
        }

        // Create a compatible bitmap
        let hbitmap = unsafe { CreateCompatibleBitmap(hdc_screen, width as i32, height as i32) };
        if hbitmap.is_invalid() {
            unsafe { let _ = DeleteDC(hdc_mem); let _ = DeleteDC(hdc_screen); }
            return Err(Error::last_os_error());
        }

        let h_old = unsafe { SelectObject(hdc_mem, hbitmap.into()) };

        // BitBlt from the virtual screen at the specified offset
        let result = unsafe {
            BitBlt(
                hdc_mem,
                0, 0,
                width as i32, height as i32,
                Some(hdc_screen),
                rect.get_left(), rect.get_top(),
                SRCCOPY,
            )
        };

        // Select the old object back and clean up
        unsafe { let _ = SelectObject(hdc_mem, h_old); }
        unsafe { let _ = DeleteDC(hdc_mem); }
        unsafe { let _ = DeleteDC(hdc_screen); }

        if result.is_err() {
            unsafe { let _ = DeleteObject(hbitmap.into()); }
            return Err(Error::last_os_error());
        }

        // Read the bitmap data
        let pixels = read_bitmap(hbitmap, width, height)?;

        // Clean up the bitmap handle
        unsafe { let _ = DeleteObject(hbitmap.into()); }

        Ok(Screenshot {
            pixels,
            width,
            height,
            format: PixelFormat::BGRA,
        })
    }

    /// Captures the screenshot of a UI element's bounding rectangle.
    ///
    /// This is a convenience method that gets the element's bounding rectangle
    /// and captures that region.
    pub fn capture_element(element: &UIElement) -> Result<Self> {
        let rect = element.get_bounding_rectangle()?;
        Self::capture_rect(rect)
    }

    /// Returns the width of the screenshot in pixels.
    pub fn width(&self) -> u32 {
        self.width
    }

    /// Returns the height of the screenshot in pixels.
    pub fn height(&self) -> u32 {
        self.height
    }

    /// Returns a reference to the raw pixel data.
    ///
    /// The pixel format is determined by [`Screenshot::format()`].
    pub fn pixels(&self) -> &[u8] {
        &self.pixels
    }

    /// Returns the pixel format of the screenshot.
    pub fn format(&self) -> PixelFormat {
        self.format
    }

    /// Converts the screenshot to top-down RGBA format.
    ///
    /// This flips the image vertically (GDI bitmaps are bottom-up) and
    /// swaps the B and R channels (BGRA -> RGBA).
    pub fn to_rgba(&self) -> Screenshot {
        let mut rgba = vec![0u8; self.pixels.len()];
        let row_bytes = (self.width as usize) * 4;

        for y in 0..self.height as usize {
            let src_row = (self.height as usize - 1 - y) * row_bytes;
            let dst_row = y * row_bytes;

            for x in 0..self.width as usize {
                let src_idx = src_row + x * 4;
                let dst_idx = dst_row + x * 4;

                // BGRA -> RGBA (swap B and R)
                rgba[dst_idx] = self.pixels[src_idx + 2];     // R
                rgba[dst_idx + 1] = self.pixels[src_idx + 1]; // G
                rgba[dst_idx + 2] = self.pixels[src_idx];     // B
                rgba[dst_idx + 3] = self.pixels[src_idx + 3]; // A
            }
        }

        Screenshot {
            pixels: rgba,
            width: self.width,
            height: self.height,
            format: PixelFormat::RGBA,
        }
    }

    /// Saves the screenshot as a BMP file.
    pub fn save_bmp(&self, path: impl AsRef<Path>) -> Result<()> {
        let bmp_data = self.to_bmp_bytes();
        std::fs::write(path, &bmp_data).map_err(|e| Error::from(format!("Failed to write BMP file: {}", e)))
    }

    /// Converts the screenshot to BMP file bytes.
    ///
    /// Returns a complete BMP file with BITMAPINFOHEADER header and pixel data.
    pub fn to_bmp_bytes(&self) -> Vec<u8> {
        let width = self.width as i32;
        let height = self.height as i32;
        let row_bytes = (self.width as usize) * 4;
        let image_size = row_bytes * self.height as usize;

        let file_size = 14 + 40 + image_size as u32; // BMP header + DIB header + pixel data

        let mut data = Vec::with_capacity(file_size as usize);

        // BMP File Header (14 bytes)
        data.extend_from_slice(b"BM");                       // Signature
        data.extend_from_slice(&file_size.to_le_bytes());    // File size
        data.extend_from_slice(&0u32.to_le_bytes());         // Reserved
        data.extend_from_slice(&54u32.to_le_bytes());        // Pixel data offset

        // BITMAPINFOHEADER (40 bytes)
        data.extend_from_slice(&40u32.to_le_bytes());        // Header size
        data.extend_from_slice(&width.to_le_bytes());        // Width
        data.extend_from_slice(&height.to_le_bytes());       // Height (positive = bottom-up)
        data.extend_from_slice(&1u16.to_le_bytes());         // Color planes
        data.extend_from_slice(&32u16.to_le_bytes());        // Bits per pixel
        data.extend_from_slice(&0u32.to_le_bytes());         // Compression (BI_RGB)
        data.extend_from_slice(&(image_size as u32).to_le_bytes()); // Image size
        data.extend_from_slice(&0i32.to_le_bytes());         // X pixels per meter
        data.extend_from_slice(&0i32.to_le_bytes());         // Y pixels per meter
        data.extend_from_slice(&0u32.to_le_bytes());         // Colors used
        data.extend_from_slice(&0u32.to_le_bytes());         // Important colors

        // Pixel data (bottom-up for BMP, which matches GDI's native format)
        data.extend_from_slice(&self.pixels);

        data
    }

    /// Saves the screenshot as a PNG file.
    ///
    /// This method converts the image to RGBA format before encoding.
    #[cfg(feature = "png")]
    pub fn save_png(&self, path: impl AsRef<Path>) -> Result<()> {
        let rgba = self.to_rgba();

        use std::fs::File;
        use std::io::BufWriter;

        let file = File::create(path)
            .map_err(|e| Error::from(format!("Failed to create PNG file: {}", e)))?;
        let writer = BufWriter::new(file);

        let mut encoder = png::Encoder::new(writer, rgba.width, rgba.height);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);

        let mut png_writer = encoder.write_header()
            .map_err(|e| Error::from(format!("Failed to write PNG header: {}", e)))?;

        png_writer.write_image_data(&rgba.pixels)
            .map_err(|e| Error::from(format!("Failed to write PNG data: {}", e)))?;

        png_writer.finish()
            .map_err(|e| Error::from(format!("Failed to finalize PNG: {}", e)))?;

        Ok(())
    }
}

/// Reads the pixel data from a GDI bitmap handle.
fn read_bitmap(hbitmap: windows::Win32::Graphics::Gdi::HBITMAP, width: u32, height: u32) -> Result<Vec<u8>> {
    // Get bitmap info to determine stride
    let mut bmp: BITMAP = BITMAP::default();
    let got = unsafe { GetObjectW(hbitmap.into(), std::mem::size_of::<BITMAP>() as i32, Some(&mut bmp as *mut _ as *mut _)) };
    if got == 0 {
        return Err(Error::last_os_error());
    }

    let stride = bmp.bmWidthBytes as usize;
    let row_bytes = (width as usize) * 4;
    let image_size = stride * height as usize;

    // Set up BITMAPINFO
    let bi = BITMAPINFOHEADER {
        biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
        biWidth: width as i32,
        biHeight: height as i32, // positive = bottom-up DIB
        biPlanes: 1,
        biBitCount: 32,
        biCompression: 0, // BI_RGB
        biSizeImage: image_size as u32,
        biXPelsPerMeter: 0,
        biYPelsPerMeter: 0,
        biClrUsed: 0,
        biClrImportant: 0,
    };

    let mut bmi = BITMAPINFO {
        bmiColors: [windows::Win32::Graphics::Gdi::RGBQUAD::default()],
        bmiHeader: bi,
    };

    let mut pixels = vec![0u8; image_size];

    let hdc = unsafe { CreateCompatibleDC(None) };
    if hdc.is_invalid() {
        return Err(Error::last_os_error());
    }

    let got = unsafe {
        GetDIBits(
            hdc,
            hbitmap,
            0,
            height,
            Some(pixels.as_mut_ptr() as *mut _),
            &mut bmi,
            DIB_RGB_COLORS,
        )
    };

    unsafe { let _ = DeleteDC(hdc); }

    if got == 0 {
        return Err(Error::last_os_error());
    }

    // If stride != row_bytes, we need to trim each row
    if stride != row_bytes {
        let mut trimmed = Vec::with_capacity(row_bytes * height as usize);
        for y in 0..height as usize {
            let start = y * stride;
            trimmed.extend_from_slice(&pixels[start..start + row_bytes]);
        }
        Ok(trimmed)
    } else {
        Ok(pixels)
    }
}

/// Gets the virtual screen rect using system metrics.
fn get_virtual_screen_rect() -> Result<Rect> {
    let x = unsafe { GetSystemMetrics(SM_XVIRTUALSCREEN) };
    let y = unsafe { GetSystemMetrics(SM_YVIRTUALSCREEN) };
    let w = unsafe { GetSystemMetrics(SM_CXVIRTUALSCREEN) };
    let h = unsafe { GetSystemMetrics(SM_CYVIRTUALSCREEN) };

    if w == 0 || h == 0 {
        return Err(Error::last_os_error());
    }

    Ok(Rect::new(x, y, x + w, y + h))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_capture_desktop() {
        let screenshot = Screenshot::capture_desktop().unwrap();
        assert!(screenshot.width() > 0);
        assert!(screenshot.height() > 0);
        assert!(!screenshot.pixels().is_empty());
        println!("Desktop screenshot: {}x{}", screenshot.width(), screenshot.height());
    }

    #[test]
    fn test_capture_rect() {
        let rect = Rect::new(0, 0, 100, 100);
        let screenshot = Screenshot::capture_rect(rect).unwrap();
        assert_eq!(screenshot.width(), 100);
        assert_eq!(screenshot.height(), 100);
    }

    #[test]
    fn test_to_rgba() {
        let rect = Rect::new(0, 0, 10, 10);
        let screenshot = Screenshot::capture_rect(rect).unwrap();
        let rgba = screenshot.to_rgba();
        assert_eq!(rgba.format(), PixelFormat::RGBA);
        assert_eq!(rgba.pixels().len(), 10 * 10 * 4);
    }

    #[test]
    fn test_bmp_bytes() {
        let rect = Rect::new(0, 0, 10, 10);
        let screenshot = Screenshot::capture_rect(rect).unwrap();
        let bmp = screenshot.to_bmp_bytes();
        assert_eq!(&bmp[0..2], b"BM");
        assert!(bmp.len() > 54);
    }
}
