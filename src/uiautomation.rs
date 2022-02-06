use std::ptr::null_mut;

use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::CLSCTX_ALL;
use windows::Win32::System::Com::COINIT_MULTITHREADED;
use windows::Win32::System::Com::CoCreateInstance;
use windows::Win32::System::Com::CoInitializeEx;
use windows::Win32::UI::Accessibility::CUIAutomation;
use windows::Win32::UI::Accessibility::IUIAutomation;
use windows::Win32::UI::Accessibility::IUIAutomationElement;
use windows::Win32::UI::Accessibility::IUIAutomationTreeWalker;

use crate::errors::Error;

pub type Result<T> = core::result::Result<T, Error>;

impl From<windows::core::Error> for Error {
    fn from(_: windows::core::Error) -> Self {
        todo!()
    }
}

#[derive(Clone)]
pub struct UIAutomation {
    automation: IUIAutomation
}

// impl Clone for UIAutomation {
//     fn clone(&self) -> Self {
//         Self { 
//             automation: self.automation.clone() 
//         }
//     }
// }

impl UIAutomation {
    pub fn new() -> Result<UIAutomation> {
        let automation: IUIAutomation;
        unsafe {
            CoInitializeEx(null_mut(), COINIT_MULTITHREADED)?;
            automation = CoCreateInstance(&CUIAutomation, None, CLSCTX_ALL)?;
        }

        Ok(UIAutomation {
            automation
        })
    }

    pub fn get_root_element(&self) -> Result<UIElement> {
        let element: IUIAutomationElement;
        unsafe {
            element = self.automation.GetRootElement()?;
        }

        Ok(UIElement::new(element))
    }

    pub fn create_tree_walker(&self) -> Result<UITreeWalker> {
        let tree_walker: IUIAutomationTreeWalker;
        unsafe {
            let condition = self.automation.CreateTrueCondition()?;
            tree_walker = self.automation.CreateTreeWalker(condition)?;
        }

        Ok(UITreeWalker::new(tree_walker))
    }

    pub fn create_matcher(&self) -> UIMatcher {
        UIMatcher {
            automation: self.clone(),
            depth: 5,
            from: None,
            condition: None
        }
    }
}

#[derive(Clone)]
pub struct UIElement {
    element: IUIAutomationElement
}

impl UIElement {
    pub fn new(element: IUIAutomationElement) -> UIElement {
        UIElement {
            element
        }
    }

    pub fn get_name(&self) -> Result<String> {
        let name: BSTR;
        unsafe {
            name = self.element.CurrentName()?;
        }

        Ok(name.to_string())
    }

    pub fn get_classname(&self) -> Result<String> {
        let classname: BSTR;
        unsafe {
            classname = self.element.CurrentClassName()?;
        }

        Ok(classname.to_string())
    }
}

#[derive(Clone)]
pub struct UITreeWalker {
    tree_walker: IUIAutomationTreeWalker
}

impl UITreeWalker {
    pub fn new(tree_walker: IUIAutomationTreeWalker) -> UITreeWalker {
        UITreeWalker {
            tree_walker
        }
    }

    pub fn get_parent(&self, element: &UIElement) -> Result<UIElement> {
        let parent: IUIAutomationElement;
        unsafe {
            parent = self.tree_walker.GetParentElement(&element.element)?;
        }

        Ok(UIElement::new(parent))
    }

    pub fn get_first_child(&self, element: &UIElement) -> Result<UIElement> {
        let child: IUIAutomationElement;
        unsafe {
            child = self.tree_walker.GetFirstChildElement(&element.element)?;
        }

        Ok(UIElement::new(child))
    }

    pub fn get_last_child(&self, element: &UIElement) -> Result<UIElement> {
        let child: IUIAutomationElement;
        unsafe {
            child = self.tree_walker.GetLastChildElement(&element.element)?;
        }

        Ok(UIElement::new(child))
    }

    pub fn get_next_sibling(&self, element: &UIElement) -> Result<UIElement> {
        let sibling: IUIAutomationElement;
        unsafe {
            sibling = self.tree_walker.GetNextSiblingElement(&element.element)?;
        }

        Ok(UIElement::new(sibling))
    }

    pub fn get_previous_sibling(&self, element: &UIElement) -> Result<UIElement> {
        let sibling: IUIAutomationElement;
        unsafe {
            sibling = self.tree_walker.GetPreviousSiblingElement(&element.element)?;
        }

        Ok(UIElement::new(sibling))
    }
}

pub trait Condition {
    fn judge(&self, element: &UIElement) -> Result<bool>;
}
pub struct UIMatcher {
    automation: UIAutomation,
    depth: u32,
    from: Option<UIElement>,
    condition: Option<Box<dyn Condition>>
}

impl UIMatcher {
    pub fn from(mut self, element: UIElement) -> Self {
        self.from = Some(element);
        self
    }

    pub fn depth(mut self, depth: u32) -> Self {
        self.depth = depth;
        self
    }
}