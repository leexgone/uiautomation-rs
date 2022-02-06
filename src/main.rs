use clap::App;
use clap::Arg;
use uiautomation_rs::uiautomation::UIAutomation;
use uiautomation_rs::uiautomation::UIElement;
use uiautomation_rs::uiautomation::UITreeWalker;
use windows::core::Result;

fn main() {
    let matches = App::new("ui automation")
        .version(env!("CARGO_PKG_VERSION"))
        .author("Steven Lee")
        .arg(Arg::with_name("print").short("p").long("print"))
        .get_matches();
    
    if matches.is_present("print") {
        print_tree();
    }

    find_by_name();
}

fn find_by_name() {
    let automation = UIAutomation::new().unwrap();
    let root = automation.get_root_element().unwrap();
    let matcher = automation.create_matcher().from(root);
}

fn print_tree() {
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