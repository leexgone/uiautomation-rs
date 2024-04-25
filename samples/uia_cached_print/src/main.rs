use uiautomation::controls::ControlType;
use uiautomation::core::UICacheRequest;
use uiautomation::types::UIProperty;
use uiautomation::variants::Variant;
use uiautomation::Result;
use uiautomation::UIAutomation;
use uiautomation::UIElement;
use uiautomation::UITreeWalker;

fn main() {
    let automation = UIAutomation::new().unwrap();
    let cache_request = automation.create_cache_request().unwrap();
    cache_request.add_property(UIProperty::ControlType).unwrap();
    cache_request.add_property(UIProperty::Name).unwrap();
    cache_request.add_property(UIProperty::ClassName).unwrap();

    let filter = automation
        .create_property_condition(
            UIProperty::ControlType,
            Variant::from(ControlType::Pane as i32),
            None,
        )
        .unwrap();

    cache_request.set_tree_filter(filter.clone()).unwrap();
    // cache_request
    //     .set_tree_scope(uiautomation::types::TreeScope::Descendants)
    //     .unwrap();

    let walker = automation.filter_tree_walker(filter).unwrap(); //automation.get_control_view_walker().unwrap();
    let root = automation
        .get_root_element_build_cache(&cache_request)
        .unwrap();

    print_element(&walker, &root, &cache_request, 0).unwrap();
}

fn print_element(
    walker: &UITreeWalker,
    element: &UIElement,
    cache_request: &UICacheRequest,
    level: usize,
) -> Result<()> {
    for _ in 0..level {
        print!(" ")
    }
    println!(
        "{} - {}",
        element.get_cached_classname()?,
        element.get_cached_name()?
    );

    if let Ok(child) = walker.get_first_child_build_cache(&element, cache_request) {
        print_element(walker, &child, cache_request, level + 1)?;

        let mut next = child;
        while let Ok(sibling) = walker.get_next_sibling_build_cache(&next, cache_request) {
            print_element(walker, &sibling, cache_request, level + 1)?;

            next = sibling;
        }
    }

    Ok(())
}
