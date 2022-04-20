#[cfg(test)]
mod tests {
    use windows::Win32::UI::Input::KeyboardAndMouse::INPUT;
    use windows::Win32::UI::Input::KeyboardAndMouse::SendInput;

    #[test]
    fn show_desktop() {
        let mut inputs: [INPUT; 4] = [INPUT::default(); 4];
    }
}