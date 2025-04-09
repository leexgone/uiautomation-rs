#[cfg(test)]
mod tests {
    use uiautomation::types::TreeScope;
    use uiautomation::types::UIProperty;
    use uiautomation::variants::Value;
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
                Variant::from(Value::ArrayI4(runtime_id)),
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
}
