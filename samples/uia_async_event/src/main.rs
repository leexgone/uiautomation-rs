use std::thread;

use uiautomation::events::CustomFocusChangedEventHandler;
use uiautomation::events::CustomStructureChangedEventHandler;
use uiautomation::events::UIFocusChangedEventHandler;
use uiautomation::events::UIStructureChangeEventHandler;
use uiautomation::processes::Process;
use uiautomation::UIAutomation;
use windows::Win32::System::Com::CoInitializeEx;

struct MyStructureChangeEventHandler {}

impl CustomStructureChangedEventHandler for MyStructureChangeEventHandler {
    fn handle(&self, sender: &uiautomation::UIElement, change_type: uiautomation::types::StructureChangeType, _runtime_id: Option<&[i32]>) -> uiautomation::Result<()> {
        println!("Structure changed: {}. Change type: {}.", sender, change_type);
        Ok(())
    }
}

struct MyFocusChangedEventHandler{}

impl CustomFocusChangedEventHandler for MyFocusChangedEventHandler {
    fn handle(&self, sender: &uiautomation::UIElement) -> uiautomation::Result<()> {
        println!("Focus changed: {}", sender);
        Ok(())
    }
}

fn main() {
    let notepad = Process::new("notepad.exe").wait_for_idle(1000).run().unwrap();
    let notepad_id = notepad.get_id();

    thread::spawn(move || {
        unsafe {
            CoInitializeEx(None, windows::Win32::System::Com::COINIT_MULTITHREADED) // init with `MULTITHREADED`, can watch event between threads
            // CoInitializeEx(None, windows::Win32::System::Com::COINIT_APARTMENTTHREADED) // init with `APARTMENTTHREADED`, can not watch event between threads
        }.unwrap();
        let automation = UIAutomation::new_direct().unwrap();  // UIAutomation::new().unwrap();
        let matcher = automation.create_matcher().timeout(5000).process_id(notepad_id).classname("Notepad");
        match matcher.find_first() {
            Ok(notepad) => {
                println!("find notepad, adding event listener...");

                let structure_changed_handler = MyStructureChangeEventHandler {};
                let structure_changed_handler = UIStructureChangeEventHandler::from(structure_changed_handler);

                automation.add_structure_changed_event_handler(&notepad, uiautomation::types::TreeScope::Subtree, None, &structure_changed_handler).unwrap();

                let focus_changed_handler = MyFocusChangedEventHandler {};
                let focus_changed_handler = UIFocusChangedEventHandler::from(focus_changed_handler);

                automation.add_focus_changed_event_handler(None, &focus_changed_handler).unwrap();

                println!("try to do something with 'notepad'...");
            },
            Err(e) => println!("failed to find notepad, {}", e),
        }
    });

    notepad.wait().unwrap();
}
