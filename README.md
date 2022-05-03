# Rust for windows uiautomation

The `uiatomation-rs` crate is a wrapper for windows uiautomation. This crate can help you make windows uiautoamtion API calls conveniently.

## Usages

Start by adding the following to your Cargo.toml file:

``` toml
[dependencies]
uiautomation = "0.0.9"
```

Make use of any windows uiautomation calls as needed.

## Examples

### Print All UIElements

``` rust
use uiautomation::Result;
use uiautomation::UIAutomation;
use uiautomation::UIElement;
use uiautomation::UITreeWalker;

fn main() {
    let automation = UIAutomation::new().unwrap();
    let walker = automation.create_tree_walker().unwrap();
    let root = automation.get_root_element().unwrap();

    print_element(&walker, &root, 0).unwrap();
}

fn print_element(walker: &UITreeWalker, element: &UIElement, level: usize) -> Result<()> {
    for _ in 0..level {
        print!(" ")
    }
    println!("{} - {}", element.get_classname()?, element.get_name()?);

    if let Ok(child) = walker.get_first_child(&element) {
        print_element(walker, &child, level + 1)?;

        let mut next = child;
        while let Ok(sibing) = walker.get_next_sibling(&next) {
            print_element(walker, &sibing, level + 1)?;

            next = sibing;
        }
    }
    
    Ok(())
}
```

### Find Notepad and Invoke File Menu

``` rust
use uiautomation::core::UIAutomation;
use uiautomation::patterns::UIInvokePattern;
use windows::Win32::UI::Accessibility::UIA_MenuItemControlTypeId;

fn main() {
    let automation = UIAutomation::new().unwrap();
    let root = automation.get_root_element().unwrap();
    let matcher = automation.create_matcher().from(root).timeout(10000).contains_name("记事本").classname("Notepad");
    if let Ok(notpad) = matcher.find_first() {
        println!("Found: {} - {}", notpad.get_name().unwrap(), notpad.get_classname().unwrap());

        let m = automation.create_matcher().from(notpad).contains_name("文件").control_type(UIA_MenuItemControlTypeId);
        if let Ok(open_menu) = m.find_first() {
            println!("Found Open: {}", open_menu.get_name().unwrap());
            let invoker: UIInvokePattern = open_menu.get_pattern().unwrap();
            invoker.invoke().unwrap();
        }
    }
}
```

### Get Properties As Variant

``` rust
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
```

### Simulate Keyboard Input

``` rust
use uiautomation::core::UIAutomation;

fn main() {
    let automation = UIAutomation::new().unwrap();
    let root = automation.get_root_element().unwrap();
    root.send_keys("{Win}D", 10).unwrap();
}
```
