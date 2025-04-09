#[cfg(test)]
mod tests {
    use uiautomation::actions::Value;
    use uiautomation::controls::ControlType;
    use uiautomation::controls::DocumentControl;
    use uiautomation::processes::Process;
    use uiautomation::types::TreeScope;
    use uiautomation::types::UIProperty;
    use uiautomation::variants::Value as VariantValue;
    use uiautomation::variants::Variant;
    use uiautomation::UIAutomation;
    use uiautomation::UIElement;

    #[test]
    fn test_runtime_id() {
        let automation = UIAutomation::new().unwrap();
        let root = automation.get_root_element().unwrap();

        // let runtime_id_var = root.get_property_value(UIProperty::RuntimeId).unwrap();
        // println!("variant = {}", runtime_id_var);

        let runtime_id = root.get_runtime_id().unwrap();
        println!("array = {:?}", runtime_id);

        // let arr = SafeArray::try_from(&runtime_id).unwrap();
        // println!("safe array = {}", arr);

        // let var = Variant::from(Value::ArrayI4(runtime_id));
        // println!("new variant = {}", var);
    
        // exit code: 0xc0000005, STATUS_ACCESS_VIOLATION occurs on next line
        let condition = automation
            .create_property_condition(
                UIProperty::RuntimeId,
                Variant::from(VariantValue::ArrayI4(runtime_id)),
                None,
            )
            .expect("Failed to create condition");
    
        let element = root
            .find_first(TreeScope::Element, &condition)
            .expect("Failed to find element");
    
        println!("{}", element.get_name().unwrap());
    }

    #[test]
    fn test_search_by_helptext() {
        let automation = UIAutomation::new().unwrap();

        let matcher = automation.create_matcher().contains_name("Microsoft Edge");  // Find Edge window
        if let Ok(edge) = matcher.find_first() {
            let matcher = automation.create_matcher().from(edge).depth(10).filter_fn(Box::new(|e: &UIElement| {
                // Search by help text. Change the search content according to the language you have localized.
                Ok(e.get_help_text()? == "搜索或输入 Web 地址") 
            }));

            if let Ok(input) = matcher.find_first() {
                println!("Found element with help text: {}", input.get_name().unwrap());
            } else {
                println!("Element not found");
            }
        }
    }

    #[test]
    fn test_send_multilines() {
        Process::create("notepad.exe").unwrap();

        let automation = UIAutomation::new().unwrap();
        let matcher = automation.create_matcher().classname("Notepad").timeout(10000);
        if let Ok(notepad) = matcher.find_first() {
            let text = "let text = r\"Customer ID: 98765
Name: Globex Corporation
Order Total: $1,234.56
Status: Pending Shipment
Contact Email: procurement@globex.example.com\"";
            // notepad.send_keys(&text, 50).unwrap();
            notepad.send_text(&text, 50).unwrap();

            let matcher = automation.create_matcher().from_ref(&notepad).control_type(ControlType::Document);
            let document = matcher.find_first().unwrap();
            let doc_ctrl = DocumentControl::try_from(document).unwrap();
            let content = doc_ctrl.get_value().unwrap();

            let raw_lines: Vec<&str> = text.split(&['\r', '\n']).collect();
            let new_lines: Vec<&str> = content.split(&['\r', '\n']).collect();
            assert_eq!(raw_lines, new_lines);
        }
    }
}
