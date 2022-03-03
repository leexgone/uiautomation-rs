use uiautomation::core::UIAutomation;

fn main() {
    let automation = UIAutomation::new().unwrap();
    let root = automation.get_root_element().unwrap();
    let matcher = automation.create_matcher().from(root).timeout(10000).contains_name("记事本");
    if let Ok(notpad) = matcher.find_first() {
        println!("Found: {} - {}", notpad.get_name().unwrap(), notpad.get_classname().unwrap());

        let m = automation.create_matcher().from(notpad).contains_name("文件");
        if let Ok(open_menu) = m.find_first() {
            println!("Found Open: {}", open_menu.get_name().unwrap());
            let invoker: UIInvokePattern = open_menu.get_pattern().unwrap();
            invoker.invoke().unwrap();
        }
    }

    if let Ok(notepads) = matcher.find_all() {
        for notepad in notepads {
            println!("Found in all: {} - {}", notepad.get_name().unwrap(), notepad.get_classname().unwrap());
            // notepad.set_focus().unwrap();
        }
    }    
}