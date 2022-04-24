#[cfg(test)]
mod tests {
    use std::mem;

    use windows::Win32::Foundation::GetLastError;
    use windows::Win32::UI::Input::KeyboardAndMouse::INPUT;
    use windows::Win32::UI::Input::KeyboardAndMouse::INPUT_0;
    use windows::Win32::UI::Input::KeyboardAndMouse::INPUT_KEYBOARD;
    use windows::Win32::UI::Input::KeyboardAndMouse::KEYBDINPUT;
    use windows::Win32::UI::Input::KeyboardAndMouse::KEYBD_EVENT_FLAGS;
    use windows::Win32::UI::Input::KeyboardAndMouse::KEYEVENTF_KEYUP;
    use windows::Win32::UI::Input::KeyboardAndMouse::SendInput;
    use windows::Win32::UI::Input::KeyboardAndMouse::VK_D;
    use windows::Win32::UI::Input::KeyboardAndMouse::VK_LWIN;

    #[test]
    fn show_desktop() {
        let inputs: [INPUT; 4] = [ INPUT {
            r#type: INPUT_KEYBOARD,
            Anonymous: INPUT_0 { 
                ki: KEYBDINPUT {
                    wVk: VK_LWIN,
                    wScan: 0,
                    dwFlags: KEYBD_EVENT_FLAGS::default(),
                    time: 0,
                    dwExtraInfo: 0
                } 
            }
        }, INPUT {
            r#type: INPUT_KEYBOARD,
            Anonymous: INPUT_0 {
                ki: KEYBDINPUT { 
                    wVk: VK_D, 
                    wScan: 0, 
                    dwFlags: KEYBD_EVENT_FLAGS::default(), 
                    time: 0, 
                    dwExtraInfo: 0 
                }
            }
        }, INPUT {
            r#type: INPUT_KEYBOARD,
            Anonymous: INPUT_0 {
                ki: KEYBDINPUT { 
                    wVk: VK_D, 
                    wScan: 0, 
                    dwFlags: KEYEVENTF_KEYUP, 
                    time: 0, 
                    dwExtraInfo: 0 
                }
            }
        }, INPUT {
            r#type: INPUT_KEYBOARD,
            Anonymous: INPUT_0 {
                ki: KEYBDINPUT { 
                    wVk: VK_LWIN, 
                    wScan: 0, 
                    dwFlags: KEYEVENTF_KEYUP, 
                    time: 0, 
                    dwExtraInfo: 0 
                }
            }
        }];

        let sent = unsafe {
            SendInput(&inputs, mem::size_of::<INPUT>() as _)
        };

        if sent == 0 {
            let err = unsafe {
                GetLastError()
            };
            println!("error code: {}", err.0);
        }
        
        assert_eq!(sent as usize, inputs.len());
    }
}