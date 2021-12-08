use windows::{
    core::*, Win32::System::Com::*, Win32::UI::Accessibility::*
};

fn main() {
    unsafe {
        CoInitializeEx(std::ptr::null_mut(), COINIT_MULTITHREADED).unwrap();

        let automation: IUIAutomation = CoCreateInstance(&CUIAutomation, None, CLSCTX_ALL).unwrap();
        let condition = automation.CreateTrueCondition().unwrap();
        let walker = automation.CreateTreeWalker(condition).unwrap();

        let root = automation.GetRootElement().unwrap(); //walker.GetFirstChildElement(None).unwrap();

        // let name = root.CurrentName().unwrap();
        // println!("{}", name);
        print_element(&walker, &root, 0).unwrap();
    }
}

unsafe fn print_element(walker: &IUIAutomationTreeWalker, element: &IUIAutomationElement, level: usize) -> Result<()> {
    for _ in 0..level {
        print!("  ");
    }
    println!("{}", element.CurrentName().unwrap());

    if let Ok(child) = walker.GetFirstChildElement(element) {
        print_element(walker, &child, level + 1)?;

        let mut next = child;
        while let Ok(sibling) = walker.GetNextSiblingElement(next) {
            print_element(walker, &sibling, level + 1)?;
            next = sibling;
        }
    }

    Ok(())
}