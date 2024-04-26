#[cfg(test)]
mod tests {
    use uiautomation::types::TreeScope;
    use uiautomation::types::UIProperty;
    use uiautomation::variants::Value;
    use uiautomation::variants::Variant;
    use uiautomation::UIAutomation;

    #[test]
    fn test_runtime_id() {
        let automation = UIAutomation::new().unwrap();
        let root = automation.get_root_element().unwrap();
        let runtime_id = root.get_runtime_id().unwrap();
    
        // exit code: 0xc0000005, STATUS_ACCESS_VIOLATION occurs on next line
        let condition = automation
            .create_property_condition(
                UIProperty::RuntimeId,
                Variant::from(Value::ArrayI4(runtime_id)),
                None,
            ).unwrap();
            // .expect("Failed to create condition");
    
        let element = root
            .find_first(TreeScope::Element, &condition)
            .expect("Failed to find element");
    
        eprintln!("{}", element.get_name().unwrap());
    }
}
