# Rust for windows uiautomation

The `uiatomation-rs` crate is a wrapper for windows uiautomation. This crate can help you make windows uiautoamtion API calls conveniently.

## Usages

Start by adding the dependency of this crate to your Cargo.toml file. Make use of any windows uiautomation calls as needed.

## Features

### Supported Features

| Feature | Description | Default |
| ------- | ----------- | ------- |
| `process` | Support process operations and filter by process id | False |
| `dialog` | Enable message boxes to show message | False |
| `input` | Support keyboard inputs | True |
| `pattern` | Support Microsoft UI Automation control patterns | - |
| `control` | Enable to wrapper ui element as control to simplify operations | True |
| `log` | Use log crate to print debug message | False |
| `all` | Enable all the above features | False |

> `pattern` is a feature that `control` depends on.

### Default Features

+ `input`, `control`, `pattern`(introduced by the `control` dependency)

> In order to be compatible with the version before `v0.19.0`, you should add `process` and `dialog` features.

## Examples

### Print All UIElements

``` rust
use uiautomation::Result;
use uiautomation::UIAutomation;
use uiautomation::UIElement;
use uiautomation::UITreeWalker;

fn main() {
    let automation = UIAutomation::new().unwrap();
    let walker = automation.get_control_view_walker().unwrap();
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
        while let Ok(sibling) = walker.get_next_sibling(&next) {
            print_element(walker, &sibling, level + 1)?;

            next = sibling;
        }
    }
    
    Ok(())
}
```

### Open Notepad and Input Text

``` rust
use uiautomation::core::UIAutomation;
use uiautomation::processes::Process;

fn main() {
    Process::create("notepad.exe").unwrap();

    let automation = UIAutomation::new().unwrap();
    let root = automation.get_root_element().unwrap();
    let matcher = automation.create_matcher().from(root).timeout(10000).classname("Notepad");
    if let Ok(notepad) = matcher.find_first() {
        println!("Found: {} - {}", notepad.get_name().unwrap(), notepad.get_classname().unwrap());

        notepad.send_keys("Hello,Rust UIAutomation!{enter}", 10).unwrap();

        let window: WindowControl = notepad.try_into().unwrap();
        window.maximize().unwrap();
    }
}
```

### Get Properties As Variant

``` rust
use uiautomation::UIAutomation;
use uiautomation::types::UIProperty;
use uiautomation::variants::Variant;

fn main() {
    let automation = UIAutomation::new().unwrap();
    let root = automation.get_root_element().unwrap();

    let name: Variant = root.get_property_value(UIProperty::Name).unwrap();
    println!("name = {}", name.get_string().unwrap());

    let ctrl_type: Variant = root.get_property_value(UIProperty::ControlType).unwrap();
    let ctrl_type_id: i32 = ctrl_type.try_into().unwrap();
    println!("control type = {}", ctrl_type_id);

    let enabled: Variant = root.get_property_value(UIProperty::IsEnabled).unwrap();
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

### Add Event Handler

``` rust
struct MyFocusChangedEventHandler{}

impl CustomFocusChangedEventHandler for MyFocusChangedEventHandler {
    fn handle(&self, sender: &uiautomation::UIElement) -> uiautomation::Result<()> {
        println!("Focus changed: {}", sender);
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
    }

    println!("waiting for notepad.exe...");
    note_proc.wait().unwrap();
}
```
