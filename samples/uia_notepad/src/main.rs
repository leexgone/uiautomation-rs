use uiautomation::actions::Window;
use uiautomation::controls::WindowControl;
use uiautomation::core::UIAutomation;
use uiautomation::inputs::Keyboard;
use uiautomation::processes::Process;

fn main() {
    Process::create("notepad.exe").unwrap();

    let automation = UIAutomation::new().unwrap();
    let root = automation.get_root_element().unwrap();
    let matcher = automation.create_matcher().from(root).timeout(10000).classname("Notepad").debug(true);
    if let Ok(notepad) = matcher.find_first() {
        println!("Found: {} - {}", notepad.get_name().unwrap(), notepad.get_classname().unwrap());

        notepad.send_keys("Hello, Rust UIAutomation!", 10).unwrap();
        notepad.send_text("\r\n{Win}D.", 10).unwrap();

        let kb = Keyboard::new().interval(10).ignore_parse_err(true);
        kb.send_keys(" {None} (Keys).").unwrap();

        notepad.hold_send_keys("{Ctrl}{Shift}", "{Left}{Left}", 50).unwrap();

        let window: WindowControl = notepad.try_into().unwrap();
        window.maximize().unwrap();
    }
}