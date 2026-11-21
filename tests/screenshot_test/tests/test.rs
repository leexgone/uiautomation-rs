use std::fs;
use std::path::PathBuf;

use uiautomation::screenshots::{PixelFormat, Screenshot};
use uiautomation::types::Rect;
use uiautomation::UIAutomation;

fn output_dir() -> PathBuf {
    let dir = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("output");
    if !dir.exists() {
        fs::create_dir_all(&dir).unwrap();
    }
    dir
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_capture_desktop() {
        let shot = Screenshot::capture_desktop().unwrap();

        assert!(shot.width() > 0, "screenshot width should be positive");
        assert!(shot.height() > 0, "screenshot height should be positive");
        assert!(!shot.pixels().is_empty(), "pixels should not be empty");

        let expected_size = (shot.width() * shot.height() * 4) as usize;
        assert_eq!(
            shot.pixels().len(),
            expected_size,
            "pixel buffer size mismatch"
        );
        assert_eq!(
            shot.format(),
            PixelFormat::BGRA,
            "default format should be BGRA"
        );

        println!("Desktop screenshot: {}x{}", shot.width(), shot.height());
    }

    #[test]
    fn test_capture_rect() {
        let rect = Rect::new(0, 0, 100, 100);
        let shot = Screenshot::capture_rect(rect).unwrap();

        assert_eq!(shot.width(), 100);
        assert_eq!(shot.height(), 100);
        assert_eq!(shot.pixels().len(), 100 * 100 * 4);
    }

    #[test]
    fn test_capture_rect_invalid() {
        let rect = Rect::new(0, 0, 0, 0);
        let result = Screenshot::capture_rect(rect);
        assert!(result.is_err(), "zero-size rect should fail");
    }

    #[test]
    fn test_to_rgba() {
        let rect = Rect::new(0, 0, 10, 10);
        let shot = Screenshot::capture_rect(rect).unwrap();
        let rgba = shot.to_rgba();

        assert_eq!(rgba.format(), PixelFormat::RGBA);
        assert_eq!(rgba.width(), 10);
        assert_eq!(rgba.height(), 10);
        assert_eq!(rgba.pixels().len(), 10 * 10 * 4);

        // Verify channels are actually swapped (BGRA -> RGBA)
        // Take the first pixel from both formats and compare
        let bgr = &shot.pixels()[0..4];
        let rgb = &rgba.pixels()[0..4];
        assert_eq!(rgb[0], bgr[2], "R and B channels should be swapped");
        assert_eq!(rgb[1], bgr[1], "G channel should be unchanged");
        assert_eq!(rgb[2], bgr[0], "B and R channels should be swapped");
        assert_eq!(rgb[3], bgr[3], "A channel should be unchanged");
    }

    #[test]
    fn test_save_bmp() {
        let rect = Rect::new(0, 0, 50, 50);
        let shot = Screenshot::capture_rect(rect).unwrap();

        let path = output_dir().join("test.bmp");
        shot.save_bmp(&path).unwrap();

        assert!(path.exists(), "BMP file should be created");

        let data = fs::read(&path).unwrap();
        assert!(data.len() > 54, "BMP file should have header + pixel data");
        assert_eq!(&data[0..2], b"BM", "BMP signature mismatch");

        // Verify BMP header fields
        let file_size = u32::from_le_bytes(data[2..6].try_into().unwrap());
        assert_eq!(file_size as usize, data.len(), "BMP file size mismatch");

        let pixel_offset = u32::from_le_bytes(data[10..14].try_into().unwrap());
        assert_eq!(pixel_offset, 54, "BMP pixel data offset should be 54");

        let bmp_width = i32::from_le_bytes(data[18..22].try_into().unwrap());
        let bmp_height = i32::from_le_bytes(data[22..26].try_into().unwrap());
        assert_eq!(bmp_width, 50, "BMP width mismatch");
        assert_eq!(bmp_height, 50, "BMP height mismatch");

        println!("BMP saved to: {}", path.display());
    }

    #[test]
    fn test_save_png() {
        let rect = Rect::new(0, 0, 50, 50);
        let shot = Screenshot::capture_rect(rect).unwrap();

        let path = output_dir().join("test.png");
        shot.save_png(&path).unwrap();

        assert!(path.exists(), "PNG file should be created");

        let data = fs::read(&path).unwrap();
        // PNG signature: 89 50 4E 47 0D 0A 1A 0A
        assert_eq!(
            &data[0..8],
            &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A],
            "PNG signature mismatch"
        );

        println!("PNG saved to: {}", path.display());
    }

    #[test]
    fn test_to_bmp_bytes() {
        let rect = Rect::new(0, 0, 32, 32);
        let shot = Screenshot::capture_rect(rect).unwrap();
        let bmp = shot.to_bmp_bytes();

        assert_eq!(&bmp[0..2], b"BM");

        let expected_size = 14 + 40 + (32 * 32 * 4) as u32;
        assert_eq!(bmp.len() as u32, expected_size, "BMP size should match");
    }

    #[test]
    fn test_element_screenshot() {
        let automation = UIAutomation::new().unwrap();
        let root = automation.get_root_element().unwrap();

        // Screenshot the desktop root element
        let shot = root.screenshot().unwrap();
        assert!(shot.width() > 0);
        assert!(shot.height() > 0);

        let path = output_dir().join("root_element.png");
        shot.save_png(&path).unwrap();
        assert!(path.exists());

        println!("Element screenshot: {}x{}", shot.width(), shot.height());
    }

    #[test]
    fn test_multi_monitor_rect() {
        // Capture a large rect that spans typical screen area
        let rect = Rect::new(-100, -100, 200, 200);
        let result = Screenshot::capture_rect(rect);
        // This should succeed on multi-monitor setups where negative coords are valid
        if let Ok(shot) = result {
            assert_eq!(shot.width(), 300);
            assert_eq!(shot.height(), 300);
        }
    }

    #[test]
    fn test_accessor_methods() {
        let shot = Screenshot::new(vec![0u8; 16], 2, 2);
        assert_eq!(shot.width(), 2);
        assert_eq!(shot.height(), 2);
        assert_eq!(shot.pixels().len(), 16);
        assert_eq!(shot.format(), PixelFormat::BGRA);
    }

    #[test]
    fn test_bmp_png_consistency() {
        // Both formats should produce valid files from the same screenshot
        let rect = Rect::new(0, 0, 20, 20);
        let shot = Screenshot::capture_rect(rect).unwrap();

        let bmp = output_dir().join("consistency.bmp");
        let png = output_dir().join("consistency.png");

        shot.save_bmp(&bmp).unwrap();
        shot.save_png(&png).unwrap();

        // Both files should exist and be non-empty
        let bmp_data = fs::read(&bmp).unwrap();
        let png_data = fs::read(&png).unwrap();

        assert!(bmp_data.len() > 54);
        assert!(png_data.len() > 8); // PNG header

        // BMP and PNG sizes will differ, but both should be valid images
        println!(
            "BMP: {} bytes, PNG: {} bytes",
            bmp_data.len(),
            png_data.len()
        );
    }
}
