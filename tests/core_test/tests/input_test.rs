#[cfg(test)]
mod tests {
    use std::env;
    use std::path::PathBuf;

    use uiautomation::inputs::Keyboard;
    use uiautomation::inputs::Mouse;
    use uiautomation::inputs::MouseButton;
    use uiautomation::UIAutomation;

    #[test]
    fn test_file_drag() {
        let desktop_path: PathBuf = env::var("USERPROFILE")
            .map(|user_profile| PathBuf::from(user_profile).join("Desktop"))
            .expect("Failed to get USERPROFILE environment variable");

        println!("Desktop path: {:?}", desktop_path);

        let test_file_path = desktop_path.join("test.txt");
        std::fs::write(&test_file_path, "This is a test file.").expect("Failed to create test.txt on Desktop");
        println!("Created file: {:?}", test_file_path);

        let automation = UIAutomation::new().unwrap();
        let matcher = automation.create_matcher().name("test.txt").timeout(5000);
        if let Ok(test_file) = matcher.find_first() {
            println!("Found file: {}", test_file);

            test_file.send_keys("{Win}D", 50).unwrap();

            let mouse = Mouse::default();
            let mut file_point = test_file.get_click_point().unwrap();
            mouse.move_to(&file_point).unwrap();

            file_point.offset(200, 50);
            mouse.drag_to(MouseButton::RIGHT, &file_point).unwrap();

            test_file.send_keys("{ESC}", 50).unwrap();

            if let Ok(recyble) = automation.create_matcher().name("回收站").find_first() {
                test_file.drag_to(&recyble).unwrap();
            } else {
                println!("Recycle Bin not found.");
            }
        }

        if automation.create_matcher().name("test.txt").timeout(2000).find_first().is_ok() {
            std::fs::remove_file(&test_file_path).expect("Failed to delete test.txt from Desktop");
            println!("Deleted file: {:?}", test_file_path);
        } else {
            println!("File 'test.txt' not found on Desktop.");
        }

        Keyboard::default().send_keys("{ALT}{TAB}").unwrap();
    }
}