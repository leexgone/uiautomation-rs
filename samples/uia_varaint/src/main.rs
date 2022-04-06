use uiautomation::UIAutomation;
use uiautomation::variants::Variant;
use windows::Win32::UI::Accessibility::UIA_ControlTypePropertyId;
use windows::Win32::UI::Accessibility::UIA_IsEnabledPropertyId;
use windows::Win32::UI::Accessibility::UIA_NamePropertyId;

fn main() {
    let automation = UIAutomation::new().unwrap();
    let root = automation.get_root_element().unwrap();

    let name: Variant = root.get_property_value(UIA_NamePropertyId).unwrap();
    println!("name = {}", name.get_string().unwrap());

    let ctrl_type: Variant = root.get_property_value(UIA_ControlTypePropertyId).unwrap();
    let ctrl_type_id: i32 = ctrl_type.try_into().unwrap();
    println!("control type = {}", ctrl_type_id);

    let enabled: Variant = root.get_property_value(UIA_IsEnabledPropertyId).unwrap();
    let enabled_str: String = enabled.try_into().unwrap();
    println!("enabled = {}", enabled_str);
}