use uiautomation::Result;
use uiautomation::UIAutomation;
use uiautomation::UIElement;
use uiautomation::UITreeWalker;
use uiautomation::controls::ControlType;
use uiautomation::types::UIProperty;
use uiautomation::variants::Variant;

fn main() {
    let automation = UIAutomation::new().unwrap();
    let filter = automation.create_property_condition(UIProperty::ControlType, Variant::from(ControlType::Pane as i32), None).unwrap();
    let walker = automation.filter_tree_walker(filter).unwrap(); //automation.get_control_view_walker().unwrap(); 
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