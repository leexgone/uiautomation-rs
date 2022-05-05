use uiautomation::core::UIAutomation;
use uiautomation::processes::Process;

fn main() {
    Process::create("notepad.exe").unwrap();

    let automation = UIAutomation::new().unwrap();
    let root = automation.get_root_element().unwrap();
    let matcher = automation.create_matcher().from(root).timeout(10000).classname("Notepad");
    if let Ok(notpad) = matcher.find_first() {
        println!("Found: {} - {}", notpad.get_name().unwrap(), notpad.get_classname().unwrap());

        notpad.send_keys("Hello,Rust UIAutomation!{enter}", 0).unwrap();

        // let m = automation.create_matcher().from(notpad).contains_name("文件").control_type(UIA_MenuItemControlTypeId);
        // if let Ok(open_menu) = m.find_first() {
        //     println!("Found Open: {}", open_menu.get_name().unwrap());
        //     let invoker: UIInvokePattern = open_menu.get_pattern().unwrap();
        //     invoker.invoke().unwrap();
        // }
    }
}