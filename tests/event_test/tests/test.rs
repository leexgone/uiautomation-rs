#[cfg(test)]
mod tests {
    use std::thread::sleep;
    use std::time::Duration;

    use event_test::MyEventHandler;
    use uiautomation::UIAutomation;
    use windows::Win32::UI::Accessibility::IUIAutomation;
    use windows::Win32::UI::Accessibility::IUIAutomationEventHandler;
    use windows::Win32::UI::Accessibility::TreeScope_Subtree;
    use windows::Win32::UI::Accessibility::UIA_Text_TextChangedEventId;

    #[test]
    fn test_event() {
        let automation = UIAutomation::new().unwrap();
        let auto_ref: &IUIAutomation = automation.as_ref();

        let matcher = automation.create_matcher().classname("Notepad");
        let notepad = matcher.find_first().unwrap();

        let handler = MyEventHandler {};
        let handler: IUIAutomationEventHandler = handler.into();

        unsafe {
            auto_ref.AddAutomationEventHandler(UIA_Text_TextChangedEventId, 
                &notepad, 
                TreeScope_Subtree, 
                None, 
                &handler).unwrap();
        }

        sleep(Duration::from_secs(60));
    }
}