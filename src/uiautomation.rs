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

#[derive(Clone)]
pub struct UIAutomation {
    automation: IUIAutomation
}

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

        Ok(UIElement::from(element))
    }

    pub fn create_tree_walker(&self) -> Result<UITreeWalker> {
        let tree_walker: IUIAutomationTreeWalker;
        unsafe {
            let condition = self.automation.CreateTrueCondition()?;
            tree_walker = self.automation.CreateTreeWalker(condition)?;
        }

        Ok(UITreeWalker::from(tree_walker))
    }

    pub fn create_matcher(&self) -> UIMatcher {
        UIMatcher::new(self.clone())
    }
}

impl From<IUIAutomation> for UIAutomation {
    fn from(automation: IUIAutomation) -> Self {
        UIAutomation {
            automation
        }
    }
}

impl Into<IUIAutomation> for UIAutomation {
    fn into(self) -> IUIAutomation {
        self.automation
    }
}

impl AsRef<IUIAutomation> for UIAutomation {
    fn as_ref(&self) -> &IUIAutomation {
        &self.automation
    }
}

#[derive(Clone)]
pub struct UIElement {
    element: IUIAutomationElement
}

impl UIElement {
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

    pub fn set_focus(&self) -> Result<()> {
        unsafe {
            self.element.SetFocus()?;
        }

        Ok(())
    }
}

impl From<IUIAutomationElement> for UIElement {
    fn from(element: IUIAutomationElement) -> Self {
        UIElement {
            element
        }
    }
}

impl Into<IUIAutomationElement> for UIElement {
    fn into(self) -> IUIAutomationElement {
        self.element
    }
}

impl AsRef<IUIAutomationElement> for UIElement {
    fn as_ref(&self) -> &IUIAutomationElement {
        &self.element
    }
}

#[derive(Clone)]
pub struct UITreeWalker {
    tree_walker: IUIAutomationTreeWalker
}

impl UITreeWalker {
    pub fn get_parent(&self, element: &UIElement) -> Result<UIElement> {
        let parent: IUIAutomationElement;
        unsafe {
            parent = self.tree_walker.GetParentElement(&element.element)?;
        }

        Ok(UIElement::from(parent))
    }

    pub fn get_first_child(&self, element: &UIElement) -> Result<UIElement> {
        let child: IUIAutomationElement;
        unsafe {
            child = self.tree_walker.GetFirstChildElement(&element.element)?;
        }

        Ok(UIElement::from(child))
    }

    pub fn get_last_child(&self, element: &UIElement) -> Result<UIElement> {
        let child: IUIAutomationElement;
        unsafe {
            child = self.tree_walker.GetLastChildElement(&element.element)?;
        }

        Ok(UIElement::from(child))
    }

    pub fn get_next_sibling(&self, element: &UIElement) -> Result<UIElement> {
        let sibling: IUIAutomationElement;
        unsafe {
            sibling = self.tree_walker.GetNextSiblingElement(&element.element)?;
        }

        Ok(UIElement::from(sibling))
    }

    pub fn get_previous_sibling(&self, element: &UIElement) -> Result<UIElement> {
        let sibling: IUIAutomationElement;
        unsafe {
            sibling = self.tree_walker.GetPreviousSiblingElement(&element.element)?;
        }

        Ok(UIElement::from(sibling))
    }
}

impl From<IUIAutomationTreeWalker> for UITreeWalker {
    fn from(tree_walker: IUIAutomationTreeWalker) -> Self {
        UITreeWalker {
            tree_walker
        }
    }
}

impl Into<IUIAutomationTreeWalker> for UITreeWalker {
    fn into(self) -> IUIAutomationTreeWalker {
        self.tree_walker
    }
}

impl AsRef<IUIAutomationTreeWalker> for UITreeWalker {
    fn as_ref(&self) -> &IUIAutomationTreeWalker {
        &self.tree_walker
    }
}

pub struct UIMatcher {
    automation: UIAutomation,
    depth: u32,
    from: Option<UIElement>,
    condition: Option<Box<dyn Condition>>
}

impl UIMatcher {
    pub fn new(automation: UIAutomation) -> Self {
        UIMatcher {
            automation,
            depth: 5,
            from: None,
            condition: None
        }
    }

    pub fn from(mut self, element: UIElement) -> Self {
        self.from = Some(element);
        self
    }

    pub fn depth(mut self, depth: u32) -> Self {
        self.depth = depth;
        self
    }

    pub fn filter(mut self, condition: Box<dyn Condition>) -> Self {
        self.condition = Some(condition);
        self
    }

    pub fn contains_name<S: Into<String>>(self, name: S) -> Self {
        let condition = NameCondition {
            value: name.into(),
            casesensitive: false,
            partial: true
        };
        self.filter(Box::new(condition))
    }

    pub fn match_name<S: Into<String>>(self, name: S) -> Self {
        let condition = NameCondition {
            value: name.into(),
            casesensitive: false,
            partial: false
        };
        self.filter(Box::new(condition))
    }

    pub fn find_first(&self) -> Result<UIElement> {
        // let root = if let Some(ref from) = self.from {
        //     from.clone()
        // } else {
        //     self.automation.get_root_element()?
        // };
        // let walker = self.automation.create_tree_walker()?;
        let (root, walker) = self.prepare()?;

        let mut elements: Vec<UIElement> = Vec::new();
        self.search(&walker, &root, &mut elements, 1, true)?;

        if elements.is_empty() {
            Err(Error::from("NOTFOUND"))
        } else {
            Ok(elements.remove(0))
        }
    }

    pub fn find_all(&self) -> Result<Vec<UIElement>> {
        let (root, walker) = self.prepare()?;

        let mut elements: Vec<UIElement> = Vec::new();
        self.search(&walker, &root, &mut elements, 1, false)?;

        if elements.is_empty() {
            Err(Error::from("NOTFOUND"))
        } else {
            Ok(elements)
        }
    }

    fn prepare(&self) -> Result<(UIElement, UITreeWalker)> {
        let root = if let Some(ref from) = self.from {
            from.clone()
        } else {
            self.automation.get_root_element()?
        };
        let walker = self.automation.create_tree_walker()?;
        
        Ok((root, walker))
    }

    fn search(&self, walker: &UITreeWalker, element: &UIElement, elements: &mut Vec<UIElement>, depth: u32, first_only: bool) -> Result<()> {
        if self.is_matched(element)? {
            elements.push(element.clone());

            if first_only {
                return Ok(());
            }
        }

        if depth < self.depth {
            let mut next = walker.get_first_child(element);
            while let Ok(ref child) = next {
                self.search(walker, child, elements, depth + 1, first_only)?;
                if first_only && !elements.is_empty() {
                    return Ok(());
                }

                next = walker.get_next_sibling(child);
            }
        }

        Ok(())
    }

    fn is_matched(&self, element: &UIElement) -> Result<bool> {
        if let Some(ref condition) = self.condition {
            condition.judge(element)
        } else {
            Ok(true)
        }
    }
}

pub trait Condition {
    fn judge(&self, element: &UIElement) -> Result<bool>;
}

struct NameCondition {
    value: String,
    casesensitive: bool,
    partial: bool
}

impl Condition for NameCondition {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        let element_name = element.get_name()?;
        let element_name = element_name.as_str();
        let condition_name = self.value.as_str();

        Ok(
            if self.partial {
                if self.casesensitive {
                    element_name.contains(condition_name)
                } else {
                    let element_name = element_name.to_lowercase();
                    let condition_name = condition_name.to_lowercase();

                    element_name.contains(&condition_name)
                }
            } else {
                if self.casesensitive {
                    element_name == condition_name
                } else {
                    element_name.eq_ignore_ascii_case(condition_name)
                }
            }
        )
    }
}