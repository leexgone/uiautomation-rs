use std::fmt::Debug;
use std::fmt::Display;
use std::ptr::null_mut;
use std::thread::sleep;
use std::time::Duration;

use chrono::Local;
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::CLSCTX_ALL;
use windows::Win32::System::Com::COINIT_MULTITHREADED;
use windows::Win32::System::Com::CoCreateInstance;
use windows::Win32::System::Com::CoInitializeEx;
use windows::Win32::UI::Accessibility::CUIAutomation;
use windows::Win32::UI::Accessibility::IUIAutomation;
use windows::Win32::UI::Accessibility::IUIAutomationElement;
use windows::Win32::UI::Accessibility::IUIAutomationElement3;
use windows::Win32::UI::Accessibility::IUIAutomationElementArray;
use windows::Win32::UI::Accessibility::IUIAutomationTreeWalker;
use windows::Win32::UI::Accessibility::OrientationType;
use windows::core::Interface;

use crate::inputs::Mouse;

use super::conditions::AndCondition;
use super::conditions::ClassNameCondition;
use super::conditions::Condition;
use super::conditions::ControlTypeCondition;
use super::conditions::NameCondition;
use super::errors::ERR_NOTFOUND;
use super::errors::ERR_TIMEOUT;
use super::errors::Error;
use super::errors::Result;
use super::inputs::Keyboard;
use super::patterns::UIPattern;
use super::types::Handle;
use super::types::Rect;
use super::types::Point;
use super::variants::Variant;

/// A wrapper for windows `IUIAutomation` interface. 
/// 
/// Exposes methods that enable Microsoft UI Automation client applications to discover, access, and filter UI Automation elements.
#[derive(Clone, Debug)]
pub struct UIAutomation {
    automation: IUIAutomation
}

impl UIAutomation {
    /// Creates a uiautomation client instance. 
    pub fn new() -> Result<UIAutomation> {
        let automation: IUIAutomation = unsafe {
            CoInitializeEx(null_mut(), COINIT_MULTITHREADED)?;
            CoCreateInstance(&CUIAutomation, None, CLSCTX_ALL)?
        };

        Ok(UIAutomation {
            automation
        })
    }

    /// Compares two UI Automation elements to determine whether they represent the same underlying UI element.
    pub fn compare_elements(&self, element1: &UIElement, element2: &UIElement) -> Result<bool> {
        let same;
        unsafe {
            same = self.automation.CompareElements(element1.as_ref(), element2.as_ref())?;
        }
        Ok(same.as_bool())
    }

    /// Retrieves a UI Automation element for the specified window.
    pub fn element_from_handle(&self, hwnd: Handle) -> Result<UIElement> {
        let element = unsafe {
            self.automation.ElementFromHandle(hwnd)?
        };

        Ok(UIElement::from(element))
    }

    /// Retrieves the UI Automation element at the specified point on the desktop.
    pub fn element_from_point(&self, point: Point) -> Result<UIElement> {
        let element = unsafe {
            self.automation.ElementFromPoint(point)?
        };

        Ok(UIElement::from(element))
    }

    /// Retrieves the UI Automation element that has the input focus.
    pub fn get_focused_element(&self) -> Result<UIElement> {
        let element = unsafe {
            self.automation.GetFocusedElement()?
        };

        Ok(UIElement::from(element))
    }

    /// Retrieves the UI Automation element that represents the desktop.
    pub fn get_root_element(&self) -> Result<UIElement> {
        let element: IUIAutomationElement;
        unsafe {
            element = self.automation.GetRootElement()?;
        }

        Ok(UIElement::from(element))
    }

    /// Retrieves a tree walker object that can be used to traverse the Microsoft UI Automation tree.
    pub fn create_tree_walker(&self) -> Result<UITreeWalker> {
        let tree_walker: IUIAutomationTreeWalker;
        unsafe {
            let condition = self.automation.CreateTrueCondition()?;
            tree_walker = self.automation.CreateTreeWalker(condition)?;
        }

        Ok(UITreeWalker::from(tree_walker))
    }

    /// Creates a UIMatcher which helps to find some UIElement.
    /// 
    /// # Examples
    /// 
    /// ```
    /// use uiautomation::UIAutomation;
    /// 
    /// let automation = UIAutomation::new().unwrap();
    /// let matcher = automation.create_matcher().depth(3).timeout(1000).classname("Start");
    /// let start_menu = matcher.find_first();
    /// assert!(start_menu.is_ok())
    /// ```
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

/// A wrapper for windows `IUIAutomationElement` interface.
/// 
/// Exposes methods and properties for a UI Automation element, which represents a UI item.
#[derive(Clone)]
pub struct UIElement {
    element: IUIAutomationElement
}

impl UIElement {
    /// Retrieves the name of the element.
    pub fn get_name(&self) -> Result<String> {
        let name: BSTR;
        unsafe {
            name = self.element.CurrentName()?;
        }

        Ok(name.to_string())
    }

    /// Retrieves the Microsoft UI Automation identifier of the element.
    pub fn get_automation_id(&self) -> Result<String> {
        let automation_id = unsafe {
            self.element.CurrentAutomationId()?
        };

        Ok(automation_id.to_string())
    }

    /// Retrieves the identifier of the process that hosts the element.
    pub fn get_process_id(&self) -> Result<i32> {
        let id = unsafe {
            self.element.CurrentProcessId()?
        };

        Ok(id)
    }

    /// Retrieves the class name of the element.
   pub fn get_classname(&self) -> Result<String> {
        let classname: BSTR;
        unsafe {
            classname = self.element.CurrentClassName()?;
        }

        Ok(classname.to_string())
    }

    /// Retrieves the control type of the element.
    pub fn get_control_type(&self) -> Result<i32> {
        let control_type = unsafe {
            self.element.CurrentControlType()?
        };
        
        Ok(control_type)
    }

    /// Retrieves a localized description of the control type of the element.
    pub fn get_localized_control_type(&self) -> Result<String> {
        let control_type = unsafe {
            self.element.CurrentLocalizedControlType()?
        };

        Ok(control_type.to_string())
    }

    /// Retrieves the accelerator key for the element.
    pub fn get_accelerator_key(&self) -> Result<String> {
        let accelerator_key = unsafe {
            self.element.CurrentAcceleratorKey()?
        };

        Ok(accelerator_key.to_string())
    }

    /// Retrieves the access key character for the element.
    pub fn get_access_key(&self) -> Result<String> {
        let access_key = unsafe {
            self.element.CurrentAccessKey()?
        };

        Ok(access_key.to_string())
    }

    /// Indicates whether the element has keyboard focus.
    pub fn has_keyboard_focus(&self) -> Result<bool> {
        let has_focus = unsafe {
            self.element.CurrentHasKeyboardFocus()?
        };

        Ok(has_focus.as_bool())
    }

    /// Indicates whether the element can accept keyboard focus.
    pub fn is_keyboard_focusable(&self) -> Result<bool> {
        let focusable = unsafe {
            self.element.CurrentIsKeyboardFocusable()?
        };

        Ok(focusable.as_bool())
    }

    /// Indicates whether the element is enabled.
    pub fn is_enabled(&self) -> Result<bool> {
        let enabled = unsafe {
            self.element.CurrentIsEnabled()?
        };

        Ok(enabled.as_bool())
    }

    /// Retrieves the help text for the element.
    pub fn get_help_text(&self) -> Result<String> {
        let text = unsafe {
            self.element.CurrentHelpText()?
        };

        Ok(text.to_string())
    }

    /// Retrieves the culture identifier for the element.
    pub fn get_culture(&self) -> Result<i32> {
        let culture = unsafe {
            self.element.CurrentCulture()?
        };

        Ok(culture)
    }

    /// Indicates whether the element is a control element.
    pub fn is_control_element(&self) -> Result<bool> {
        let is_control = unsafe {
            self.element.CurrentIsControlElement()?
        };

        Ok(is_control.as_bool())
    }

    /// Indicates whether the element is a content element.
    pub fn is_content_element(&self) -> Result<bool> {
        let is_content = unsafe {
            self.element.CurrentIsContentElement()?
        };

        Ok(is_content.as_bool())
    }

    /// Indicates whether the element contains a disguised password.
    pub fn is_password(&self) -> Result<bool> {
        let is_password = unsafe {
            self.element.CurrentIsPassword()?
        };

        Ok(is_password.as_bool())
    }

    /// Retrieves the window handle of the element.
    pub fn get_native_window_handle(&self) -> Result<Handle> {
        let handle = unsafe {
            self.element.CurrentNativeWindowHandle()?
        };

        Ok(handle.into())
    }

    /// Retrieves a description of the type of UI item represented by the element.
    pub fn get_item_type(&self) -> Result<String> {
        let item_type = unsafe {
            self.element.CurrentItemType()?
        };

        Ok(item_type.to_string())
    }

    /// Indicates whether the element is off-screen.
    pub fn is_off_screen(&self) -> Result<bool> {
        let off_screen = unsafe {
            self.element.CurrentIsOffscreen()?
        };

        Ok(off_screen.as_bool())
    }

    /// Retrieves a value that indicates the orientation of the element.
    pub fn get_orientation(&self) -> Result<OrientationType> {
        let orientation = unsafe {
            self.element.CurrentOrientation()?
        };

        Ok(orientation)
    }

    /// Retrieves the name of the underlying UI framework.
    pub fn get_framework_id(&self) -> Result<String> {
        let id = unsafe {
            self.element.CurrentFrameworkId()?
        };

        Ok(id.to_string())
    }

    /// Indicates whether the element is required to be filled out on a form.
    pub fn is_required_for_form(&self) -> Result<bool> {
        let required = unsafe {
            self.element.CurrentIsRequiredForForm()?
        };

        Ok(required.as_bool())
    }

    /// Indicates whether the element contains valid data for a form.
    pub fn is_data_valid_for_form(&self) -> Result<bool> {
        let valid = unsafe {
            self.element.CurrentIsDataValidForForm()?
        };

        Ok(valid.as_bool())
    }

    /// Retrieves the description of the status of an item in an element.
    pub fn get_item_status(&self) -> Result<String> {
        let status = unsafe {
            self.element.CurrentItemStatus()?
        };

        Ok(status.to_string())
    }

    /// Retrieves the coordinates of the rectangle that completely encloses the element.
    pub fn get_bounding_rectangle(&self) -> Result<Rect> {
        let rect = unsafe {
            self.element.CurrentBoundingRectangle()?
        };

        Ok(rect.into())
    }

    /// Retrieves the element that contains the text label for this element.
    pub fn get_labeled_by(&self) -> Result<UIElement> {
        let labeled_by = unsafe {
            self.element.CurrentLabeledBy()?
        };

        Ok(UIElement::from(labeled_by))
    }

    /// Retrieves an array of elements for which this element serves as the controller.
    pub fn get_controller_for(&self) -> Result<Vec<UIElement>> {
        let elements = unsafe {
            self.element.CurrentControllerFor()?
        };

        Self::to_elements(elements)
    }

    /// Retrieves an array of elements that describe this element.
    pub fn get_described_by(&self) -> Result<Vec<UIElement>> {
        let elements = unsafe {
            self.element.CurrentDescribedBy()?
        };

        Self::to_elements(elements)
    }

    /// Retrieves an array of elements that indicates the reading order after the current element.
    pub fn get_flows_to(&self) -> Result<Vec<UIElement>> {
        let elements = unsafe {
            self.element.CurrentFlowsTo()?
        };

        Self::to_elements(elements)
    }

    /// Retrieves a description of the provider for this element.
    pub fn get_provider_description(&self) -> Result<String> {
        let descr = unsafe {
            self.element.CurrentProviderDescription()?
        };

        Ok(descr.to_string())
    }

    /// Sets the keyboard focus to this UI Automation element.
    pub fn set_focus(&self) -> Result<()> {
        unsafe {
            self.element.SetFocus()?;
        }

        Ok(())
    }

    /// Retrieves the control pattern interface of the specified pattern `<T>` from this UI Automation element.
    pub fn get_pattern<T: UIPattern>(&self) -> Result<T> {
        let pattern = unsafe {
            self.element.GetCurrentPattern(T::pattern_id())?
        };

        T::new(pattern)
    }

    /// Retrieves the current value of a property for this UI Automation element.
    pub fn get_property_value(&self, property_id: i32) -> Result<Variant> {
        let value = unsafe {
            self.element.GetCurrentPropertyValue(property_id)?
        };

        Ok(value.into())
    }

    /// Programmatically invokes a context menu on the target element.
    pub fn show_context_menu(&self) -> Result<()> {
        let element3: IUIAutomationElement3 = self.element.cast()?;
        unsafe {
            element3.ShowContextMenu()?
        }
        
        Ok(())
    }

    pub(crate) fn to_elements(elements: IUIAutomationElementArray) -> Result<Vec<UIElement>> {
        let mut arr: Vec<UIElement> = Vec::new();
        unsafe {
            for i in 0..elements.Length()? {
                let elem = UIElement::from(elements.GetElement(i)?);
                arr.push(elem);
            }
        }

        Ok(arr)
    }

    /// Simulate typing `keys` on keyboard.
    /// 
    /// `{}` is used for some special keys. For example: `{ctrl}{alt}{delete}`, `{shift}{home}`.
    /// 
    /// `()` is used for group keys. For example: `{ctrl}(AB)` types `Ctrl+A+B`.
    /// 
    /// `{}()` can be quoted by `{}`. For example: `{{}Hi,{(}rust!{)}{}}` types `{Hi,(rust)}`.
    /// 
    /// `interval` is the milliseconds between keys. `0` is the default value.
    /// 
    /// # Examples
    /// 
    /// ```
    /// use uiautomation::core::UIAutomation;
    /// 
    /// let automation = UIAutomation::new().unwrap();
    /// let root = automation.get_root_element().unwrap();
    /// root.send_keys("{Win}D", 0).unwrap();
    /// ```
    pub fn send_keys(&self, keys: &str, interval: u64) -> Result<()> {
        self.set_focus()?;
        
        let kb = Keyboard::new();
        kb.interval(interval).send_keys(keys)
    }

    /// Simulates mouse left click event on the element.
    pub fn click(&self) -> Result<()> {
        self.set_focus()?;

        let point = self.get_click_point()?;
        let mouse = Mouse::default();
        mouse.click(point)
    }

    /// Simulates mouse double click event on the element.
    pub fn double_click(&self) -> Result<()> {
        self.set_focus()?;
        
        let point = self.get_click_point()?;
        let mouse = Mouse::default();
        mouse.double_click(point)
    }

    /// Simulates mouse right click event on the element.
    pub fn right_click(&self) -> Result<()> {
        self.set_focus()?;

        let point = self.get_click_point()?;
        let mouse = Mouse::default();
        mouse.right_click(point)
    }

    fn get_click_point(&self) -> Result<Point> {
        let rect = self.get_bounding_rectangle()?;
        let point = Point::new((rect.get_left() + rect.get_right()) / 2, (rect.get_top() + rect.get_bottom()) / 2);
        Ok(point)
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

impl Display for UIElement {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let name = self.get_name().unwrap_or(String::from("(NONE)"));
        let control_type = self.get_localized_control_type().unwrap_or(String::from("UNKNOWN_TYPE"));

        write!(f, "{} {}", name, control_type)
    }
}

impl Debug for UIElement {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut d = f.debug_struct("UIElement");

        if let Ok(name) = self.get_name() {
            d.field("name", &name);
        };
        if let Ok(control_type) = self.get_control_type() {
            d.field("control_type", &control_type);
        };
        if let Ok(classname) = self.get_classname() {
            d.field("classname", &classname);
        };

        d.finish()
    }
}

/// A wrapper for windows `IUIAutomationTreeWalker` interface.
/// 
/// Exposes properties and methods that UI Automation client applications use to view and navigate the UI Automation elements on the desktop.
#[derive(Clone)]
pub struct UITreeWalker {
    tree_walker: IUIAutomationTreeWalker
}

impl UITreeWalker {
    /// Retrieves the parent element of the specified UI Automation element.
    pub fn get_parent(&self, element: &UIElement) -> Result<UIElement> {
        let parent: IUIAutomationElement;
        unsafe {
            parent = self.tree_walker.GetParentElement(&element.element)?;
        }

        Ok(UIElement::from(parent))
    }

    /// Retrieves the first child element of the specified UI Automation element.
    pub fn get_first_child(&self, element: &UIElement) -> Result<UIElement> {
        let child: IUIAutomationElement;
        unsafe {
            child = self.tree_walker.GetFirstChildElement(&element.element)?;
        }

        Ok(UIElement::from(child))
    }

    /// Retrieves the last child element of the specified UI Automation element.
    pub fn get_last_child(&self, element: &UIElement) -> Result<UIElement> {
        let child: IUIAutomationElement;
        unsafe {
            child = self.tree_walker.GetLastChildElement(&element.element)?;
        }

        Ok(UIElement::from(child))
    }

    /// Retrieves the next sibling element of the specified UI Automation element.
    pub fn get_next_sibling(&self, element: &UIElement) -> Result<UIElement> {
        let sibling: IUIAutomationElement;
        unsafe {
            sibling = self.tree_walker.GetNextSiblingElement(&element.element)?;
        }

        Ok(UIElement::from(sibling))
    }

    /// Retrieves the previous sibling element of the specified UI Automation element.
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

/// Defines filter conditions to match specific UI Element.
/// 
/// `UIMatcher` can find first element or find all elements.
pub struct UIMatcher {
    automation: UIAutomation,
    depth: u32,
    from: Option<UIElement>,
    condition: Option<Box<dyn Condition>>,
    timeout: u64,
    interval: u64,
    debug: bool
}

impl UIMatcher {
    /// Creates a matcher with `automation`.
    pub fn new(automation: UIAutomation) -> Self {
        UIMatcher {
            automation,
            depth: 7,
            from: None,
            condition: None,
            timeout: 3000,
            interval: 100,
            debug: false
        }
    }

    /// Sets the root element of the UIAutomation tree whitch should be searched from.
    /// 
    /// The root element is desktop by default.
    pub fn from(mut self, element: UIElement) -> Self {
        self.from = Some(element);
        self
    }

    /// Sets the depth of the search path. The default depth is `7`.
    pub fn depth(mut self, depth: u32) -> Self {
        self.depth = depth;
        self
    }

    /// Sets the the time in millionseconds for matching element. The default timeout is 3000 millionseconds(3 seconds).
    /// 
    /// A timeout error will occur after this time.
    pub fn timeout(mut self, timeout: u64) -> Self {
        self.timeout = timeout;
        self
    }

    /// Set the interval time in millionseconds for retrying. The default interval time is 100 millionseconds.
    pub fn interval(mut self, interval: u64) -> Self {
        self.interval = interval;
        self
    }

    /// Appends a filter condition which is used as `and` logic.
    pub fn filter(mut self, condition: Box<dyn Condition>) -> Self {
        let filter = if let Some(raw) = self.condition {
            Box::new(AndCondition::new(raw, condition))
        } else {
            condition
        };
        self.condition = Some(filter);
        self
    }

    /// Append a filter whitch match specific casesensitive name.
    pub fn name<S: Into<String>>(self, name: S) -> Self {
        let condition = NameCondition {
            value: name.into(),
            casesensitive: true,
            partial: false
        };

        self.filter(Box::new(condition))
    }

    /// Append a filter whitch name contains specific text (ignore casesensitive).
    pub fn contains_name<S: Into<String>>(self, name: S) -> Self {
        let condition = NameCondition {
            value: name.into(),
            casesensitive: false,
            partial: true
        };
        self.filter(Box::new(condition))
    }

    /// Append a filter whitch matches specific name (ignore casesensitive).
    pub fn match_name<S: Into<String>>(self, name: S) -> Self {
        let condition = NameCondition {
            value: name.into(),
            casesensitive: false,
            partial: false
        };
        self.filter(Box::new(condition))
    }

    /// Filters by classname.
    pub fn classname<S: Into<String>>(self, classname: S) -> Self {
        let condition = ClassNameCondition {
            classname: classname.into()
        };
        self.filter(Box::new(condition))        
    }

    /// Filters by control type.
    pub fn control_type(self, control_type: i32) -> Self {
        let condition = ControlTypeCondition {
            control_type
        };
        self.filter(Box::new(condition))
    }

    /// Set `debug` as `true` to enable debug mode. The debug mode is `false` by default.
    pub fn debug(mut self, debug: bool) -> Self {
        self.debug = debug;
        self
    }

    /// Finds first element.
    pub fn find_first(&self) -> Result<UIElement> {
        let elements = self.find(true)?;

        if elements.is_empty() {
            Err(Error::new(ERR_NOTFOUND, "can not find element"))
        } else {
            Ok(elements[0].clone())
        }
    }

    /// Finds all elements.
    pub fn find_all(&self) -> Result<Vec<UIElement>> {
        let elements = self.find(false)?;

        if elements.is_empty() {
            Err(Error::new(ERR_NOTFOUND, "can not find element"))
        } else {
            Ok(elements)
        }
    }

    fn find(&self, first_only: bool) -> Result<Vec<UIElement>> {
        let mut elements: Vec<UIElement> = Vec::new();
        let start = Local::now().timestamp_millis();
        loop {
            if self.debug {
                println!("Try to match element...")
            }
            
            let (root, walker) = self.prepare()?;
            self.search(&walker, &root, &mut elements, 1, first_only)?;

            if !elements.is_empty() || self.timeout <= 0 {
                break;
            }

            let now = Local::now().timestamp_millis();
            if now - start >= self.timeout as i64 {
                return Err(Error::new(ERR_TIMEOUT, "find time out"));
            }

            sleep(Duration::from_millis(self.interval));
        } 

        Ok(elements)
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
        let ret = if let Some(ref condition) = self.condition {
            condition.judge(element)?
        } else {
            true
        };

        if self.debug {
            println!("{:?} -> {}", element, ret);
        }

        Ok(ret)
    }
}

#[cfg(test)]
mod tests {
    use windows::Win32::UI::Accessibility::UIA_MenuItemControlTypeId;

    use crate::UIAutomation;

    #[test]
    fn test_zh_input() {
        let automation = UIAutomation::new().unwrap();
        let matcher = automation.create_matcher().depth(2).classname("Notepad").timeout(1000);
        if let Ok(notepad) = matcher.find_first() {
            notepad.send_keys("你好！{enter}", 0).unwrap();
            notepad.send_keys("Hello!", 0).unwrap();
        }
    }

    #[test]
    fn test_menu_click() {
        let automation = UIAutomation::new().unwrap();
        let matcher = automation.create_matcher().depth(2).classname("Notepad").timeout(1000);
        if let Ok(notepad) = matcher.find_first() {
            let matcher = automation.create_matcher().control_type(UIA_MenuItemControlTypeId).from(notepad.clone()).name("文件").depth(5).timeout(1000);
            if let Ok(menu_item) = matcher.find_first() {
                menu_item.click().unwrap();
            }
        }
    }
}
