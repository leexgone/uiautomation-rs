use uiautomation::events::CustomFocusChangedEventHandler;
use uiautomation::events::CustomPropertyChangedEventHandlerFn;
use uiautomation::events::CustomStructureChangedEventHandler;
use uiautomation::events::UIFocusChangedEventHandler;
use uiautomation::events::UIPropertyChangedEventHandler;
use uiautomation::events::UIStructureChangeEventHandler;
use uiautomation::processes::Process;
use uiautomation::types::UIProperty;
use uiautomation::UIAutomation;

struct MyFocusChangedEventHandler{}

impl CustomFocusChangedEventHandler for MyFocusChangedEventHandler {
    fn handle(&self, sender: &uiautomation::UIElement) -> uiautomation::Result<()> {
        println!("Focus changed: {}", sender);
        Ok(())
    }
}

struct MyStructureChangeEventHandler {}

impl CustomStructureChangedEventHandler for MyStructureChangeEventHandler {
    fn handle(&self, sender: &uiautomation::UIElement, change_type: uiautomation::types::StructureChangeType, _runtime_id: Option<&[i32]>) -> uiautomation::Result<()> {
        println!("Structure changed: {}. Change type: {}.", sender, change_type);
        Ok(())
    }
}

fn main() {
    let note_proc = Process::create("notepad.exe").unwrap();

    let automation = UIAutomation::new().unwrap();
    let root = automation.get_root_element().unwrap();
    let matcher = automation.create_matcher().from(root).timeout(10000).classname("Notepad");
    if let Ok(notepad) = matcher.find_first() {
        let focus_changed_handler = MyFocusChangedEventHandler {};
        let focus_changed_handler = UIFocusChangedEventHandler::from(focus_changed_handler);

        automation.add_focus_changed_event_handler(None, &focus_changed_handler).unwrap();

        let text_changed_handler: Box<CustomPropertyChangedEventHandlerFn> = Box::new(|sender, property, value| {
            println!("Property changed: {}.{:?} = {}", sender, property, value);
            Ok(())
        });
        let text_changed_handler = UIPropertyChangedEventHandler::from(text_changed_handler);

        automation.add_property_changed_event_handler(&notepad, uiautomation::types::TreeScope::Subtree, None, &text_changed_handler, &[UIProperty::ValueValue]).unwrap();

        let structure_changed_handler = MyStructureChangeEventHandler {};
        let structure_changed_handler = UIStructureChangeEventHandler::from(structure_changed_handler);

        automation.add_structure_changed_event_handler(&notepad, uiautomation::types::TreeScope::Subtree, None, &structure_changed_handler).unwrap();
    }

    println!("waiting for notepad.exe...");
    note_proc.wait().unwrap();
}
