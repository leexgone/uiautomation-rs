use std::fmt::Debug;
use std::fmt::Display;
use std::thread::sleep;
use std::time::Duration;

use chrono::Local;
use windows::core::Param;
use windows::Win32::System::Com::CLSCTX_ALL;
use windows::Win32::System::Com::COINIT_MULTITHREADED;
use windows::Win32::System::Com::CoCreateInstance;
use windows::Win32::System::Com::CoInitializeEx;
use windows::Win32::UI::Accessibility::CUIAutomation;
use windows::Win32::UI::Accessibility::IUIAutomation;
use windows::Win32::UI::Accessibility::IUIAutomationAndCondition;
use windows::Win32::UI::Accessibility::IUIAutomationBoolCondition;
use windows::Win32::UI::Accessibility::IUIAutomationCacheRequest;
use windows::Win32::UI::Accessibility::IUIAutomationCondition;
use windows::Win32::UI::Accessibility::IUIAutomationElement;
use windows::Win32::UI::Accessibility::IUIAutomationElement3;
use windows::Win32::UI::Accessibility::IUIAutomationElement8;
use windows::Win32::UI::Accessibility::IUIAutomationElement9;
use windows::Win32::UI::Accessibility::IUIAutomationElementArray;
use windows::Win32::UI::Accessibility::IUIAutomationNotCondition;
use windows::Win32::UI::Accessibility::IUIAutomationOrCondition;
use windows::Win32::UI::Accessibility::IUIAutomationPropertyCondition;
use windows::Win32::UI::Accessibility::IUIAutomationTreeWalker;
use windows::core::IUnknown;
use windows::core::Interface;
use windows::Win32::UI::Accessibility::UIA_PROPERTY_ID;

use crate::controls::ControlType;
use crate::events::UIEventHandler;
use crate::events::UIEventType;
use crate::events::UIFocusChangedEventHandler;
use crate::events::UIPropertyChangedEventHandler;
use crate::events::UIStructureChangeEventHandler;
use crate::filters::FnFilter;
use crate::inputs::Mouse;
use crate::patterns::UIPatternType;
use crate::types::ElementMode;
use crate::types::HeadingLevel;
use crate::types::OrientationType;
use crate::types::PropertyConditionFlags;
use crate::types::TreeScope;
use crate::types::UIProperty;
use crate::variants::SafeArray;

use super::filters::ClassNameFilter;
use super::filters::MatcherFilter;
use super::filters::ControlTypeFilter;
use super::filters::NameFilter;
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
#[derive(Debug, Clone)]
pub struct UIAutomation {
    automation: IUIAutomation
}

impl UIAutomation {
    /// Creates a uiautomation client instance.
    ///
    /// This method initializes the COM library each time, sets the thread's concurrency model as `COINIT_MULTITHREADED`.
    pub fn new() -> Result<UIAutomation> {
        let result = unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED)
        };

        if result.is_ok() {
            UIAutomation::new_direct()
        } else {
            Err(result.into())
        }
    }

    /// Creates a uiautomation client instance without initializing the COM library.
    ///
    /// The COM library should be initialized manually before invoking.
    pub fn new_direct() -> Result<UIAutomation> {
        let automation: IUIAutomation = unsafe {
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

    /// Retrieves a UI Automation element for the specified window, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    pub fn element_from_handle_build_cache(&self, hwnd: Handle, cache_request: &UICacheRequest) -> Result<UIElement> {
        let element = unsafe {
            self.automation.ElementFromHandleBuildCache(hwnd, cache_request)?
        };
        Ok(element.into())
    }

    /// Retrieves the UI Automation element at the specified point on the desktop.
    pub fn element_from_point(&self, point: Point) -> Result<UIElement> {
        let element = unsafe {
            self.automation.ElementFromPoint(point.into())?
        };

        Ok(UIElement::from(element))
    }

    /// Retrieves the UI Automation element at the specified point on the desktop, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    pub fn element_from_point_build_cache(&self, point: Point, cache_request: &UICacheRequest) -> Result<UIElement> {
        let element = unsafe {
            self.automation.ElementFromPointBuildCache(point.into(), cache_request)?
        };
        Ok(element.into())
    }

    /// Retrieves the UI Automation element that has the input focus.
    pub fn get_focused_element(&self) -> Result<UIElement> {
        let element = unsafe {
            self.automation.GetFocusedElement()?
        };

        Ok(UIElement::from(element))
    }

    /// Retrieves the UI Automation element that has the input focus, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    pub fn get_focused_element_build_cache(&self, cache_request: &UICacheRequest) -> Result<UIElement> {
        let element = unsafe {
            self.automation.GetFocusedElementBuildCache(cache_request)?
        };
        Ok(element.into())
    }

    /// Retrieves the UI Automation element that represents the desktop.
    pub fn get_root_element(&self) -> Result<UIElement> {
        let element = unsafe {
            self.automation.GetRootElement()?
        };

        Ok(UIElement::from(element))
    }

    /// Retrieves the UI Automation element that represents the desktop, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    pub fn get_root_element_build_cache(&self, cache_request: &UICacheRequest) -> Result<UIElement> {
        let element = unsafe {
            self.automation.GetRootElementBuildCache(cache_request)?
        };
        Ok(element.into())
    }

    /// Retrieves a tree walker object that can be used to traverse the Microsoft UI Automation tree.
    ///
    /// # Examples
    ///
    /// ```
    /// use uiautomation::UIAutomation;
    ///
    /// let automation = UIAutomation::new().unwrap();
    /// let root = automation.get_root_element().unwrap();
    /// let walker = automation.create_tree_walker().unwrap();
    /// let child = walker.get_first_child(&root);
    /// assert!(child.is_ok());
    /// ```
    pub fn create_tree_walker(&self) -> Result<UITreeWalker> {
        let tree_walker = unsafe {
            let condition = self.create_true_condition()?; //self.automation.CreateTrueCondition()?;
            self.automation.CreateTreeWalker(condition)?
        };

        Ok(UITreeWalker::from(tree_walker))
    }

    /// Retrieves a filtered tree walker object that can be used to traverse the Microsoft UI Automation tree.
    pub fn filter_tree_walker(&self, condition: UICondition) -> Result<UITreeWalker> {
        let tree_walker = unsafe {
            self.automation.CreateTreeWalker(condition)?
        };
        Ok(tree_walker.into())
    }

    /// Retrieves a tree walker object used to traverse an unfiltered view of the Microsoft UI Automation tree.
    pub fn get_raw_view_walker(&self) -> Result<UITreeWalker> {
        let walker = unsafe {
            self.automation.RawViewWalker()?
        };
        Ok(walker.into())
    }

    /// Retrieves a predefined UICondition that selects control elements.
    pub fn get_control_view_condition(&self) -> Result<UICondition> {
        let condition = unsafe {
            self.automation.ControlViewCondition()?
        };
        Ok(condition.into())
    }

    /// Retrieves a predefined UITreeWalker interface that selects control elements.
    pub fn get_control_view_walker(&self) -> Result<UITreeWalker> {
        let walker = unsafe {
            self.automation.ControlViewWalker()?
        };
        Ok(walker.into())
    }

    /// Retrieves a predefined UICondition that selects content elements.
    pub fn get_content_view_condition(&self) -> Result<UICondition> {
        let condition = unsafe {
            self.automation.ContentViewCondition()?
        };
        Ok(condition.into())
    }

    /// Retrieves a UITreeWalker interface used to discover content elements.
    pub fn get_content_view_walker(&self) -> Result<UITreeWalker> {
        let walker = unsafe {
            self.automation.ContentViewWalker()?
        };
        Ok(walker.into())
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
    /// if let Ok(start_menu) = matcher.find_first() {
    ///     println!("Found startmenu!")
    /// }
    /// ```
    pub fn create_matcher(&self) -> UIMatcher {
        UIMatcher::new(self.clone())
    }

    /// Retrieves a predefined condition that selects all elements.
    pub fn create_true_condition(&self) -> Result<UICondition> {
        let condition = unsafe {
            self.automation.CreateTrueCondition()?
        };
        Ok(condition.into())
    }

    /// Creates a condition that is always false.
    pub fn create_false_condition(&self) -> Result<UICondition> {
        let condition = unsafe {
            self.automation.CreateFalseCondition()?
        };
        Ok(condition.into())
    }

    /// Creates a condition that is the negative of a specified condition.
    pub fn create_not_condition(&self, condition: UICondition) -> Result<UICondition> {
        // let condition: IUIAutomationCondition = condition.into();
        let result = unsafe {
            self.automation.CreateNotCondition(condition)?
        };
        Ok(result.into())
    }

    /// Creates a condition that selects elements that match both of two conditions.
    pub fn create_and_condition(&self, condition1: UICondition, condition2: UICondition) -> Result<UICondition> {
        let result = unsafe {
            self.automation.CreateAndCondition(condition1, condition2)?
        };
        Ok(result.into())
    }

    /// Creates a combination of two conditions where a match exists if either of the conditions is true.
    pub fn create_or_condition(&self, condition1: UICondition, condition2: UICondition) -> Result<UICondition> {
        let result = unsafe {
            self.automation.CreateOrCondition(condition1, condition2)?
        };
        Ok(result.into())
    }

    /// Creates a condition that selects elements that have a property with the specified value, using optional flags.
    pub fn create_property_condition(&self, property: UIProperty, value: Variant, flags: Option<PropertyConditionFlags>) -> Result<UICondition> {
        let condition = unsafe {
            if let Some(flags) = flags {
                self.automation.CreatePropertyConditionEx(property.into(), value, flags.into())?
            } else {
                self.automation.CreatePropertyCondition(property.into(), value)?
            }
        };
        Ok(condition.into())
    }

    /// Creates a UICacheRequest object that specifies the properties and control patterns to be cached for an element.
    pub fn create_cache_request(&self) -> Result<UICacheRequest> {
        let request = unsafe { self.automation.CreateCacheRequest()? };
        Ok(request.into())
    }

    /// Registers a method that handles Microsoft UI Automation events.
    pub fn add_automation_event_handler(&self, event_type: UIEventType, element: &UIElement, scope: TreeScope, cache_request: Option<&UICacheRequest>, handler: &UIEventHandler) -> Result<()> {
        let cache_request = cache_request.map(|r| r.as_ref());
        unsafe {
            self.automation.AddAutomationEventHandler(event_type.into(), element, scope.into(), cache_request, handler)?
        };
        Ok(())
    }

    /// Removes the specified UI Automation event handler.
    pub fn remove_automation_event_handler(&self, event_type: UIEventType, element: &UIElement, handler: &UIEventHandler) -> Result<()> {
        unsafe {
            self.automation.RemoveAutomationEventHandler(event_type.into(), element, handler)?
        };
        Ok(())
    }

    /// Registers a method that handles and array of property-changed events.
    pub fn add_property_changed_event_handler(&self, element: &UIElement, scope: TreeScope, cache_request: Option<&UICacheRequest>, handler: &UIPropertyChangedEventHandler, properties: &[UIProperty]) -> Result<()> {
        let cache_request = cache_request.map(|r| r.as_ref());
        let prop_arr: Vec<UIA_PROPERTY_ID> = properties.iter().map(|p| (*p).into()).collect();
        unsafe {
            self.automation.AddPropertyChangedEventHandlerNativeArray(element, scope.into(), cache_request, handler, &prop_arr)?
        };
        Ok(())
    }

    /// Removes a property-changed event handler.
    pub fn remove_property_changed_event_handler(&self, element: &UIElement, handler: &UIPropertyChangedEventHandler) -> Result<()> {
        unsafe {
            self.automation.RemovePropertyChangedEventHandler(element, handler)?
        };
        Ok(())
    }

    /// Registers a method that handles structure-changed events.
    pub fn add_structure_changed_event_handler(&self, element: &UIElement, scope: TreeScope, cache_request: Option<&UICacheRequest>, handler: &UIStructureChangeEventHandler) -> Result<()> {
        let cache_request = cache_request.map(|r| r.as_ref());
        unsafe {
            self.automation.AddStructureChangedEventHandler(element, scope.into(), cache_request, handler)?
        };
        Ok(())
    }

    /// Removes a structure-changed event handler.
    pub fn remove_structure_changed_event_handler(&self, element: &UIElement, handler: &UIStructureChangeEventHandler) -> Result<()> {
        unsafe {
            self.automation.RemoveStructureChangedEventHandler(element, handler)?
        };
        Ok(())
    }

    /// Registers a method that handles focus-changed events.
    pub fn add_focus_changed_event_handler(&self, cache_request: Option<&UICacheRequest>, handler: &UIFocusChangedEventHandler) -> Result<()> {
        let cache_request = cache_request.map(|r| r.as_ref());
        unsafe {
            self.automation.AddFocusChangedEventHandler(cache_request, handler)?
        };
        Ok(())
    }

    /// Removes a focus-changed event handler.
    pub fn remove_focus_changed_event_handler(&self, handler: &UIFocusChangedEventHandler) -> Result<()> {
        unsafe {
            self.automation.RemoveFocusChangedEventHandler(handler)?
        };
        Ok(())
    }

    /// Removes all registered Microsoft UI Automation event handlers.
    pub fn remove_all_event_handlers(&self) -> Result<()> {
        unsafe {
            self.automation.RemoveAllEventHandlers()?
        };
        Ok(())
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
    /// Retrieves a new UI Automation element with an updated cache.
    pub fn build_updated_cache(&self, cache_request: &UICacheRequest) -> Result<UIElement> {
        let element = unsafe {
            self.element.BuildUpdatedCache(cache_request)?
        };
        Ok(element.into())
    }

    /// Retrieves the first child or descendant element that matches the specified condition.
    pub fn find_first(&self, scope: TreeScope, condition: &UICondition) -> Result<UIElement> {
        let result = unsafe {
            self.element.FindFirst(scope.into(), condition.as_ref())?
        };
        Ok(result.into())
    }

    /// Retrieves the first child or descendant element that matches the specified condition, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    pub fn find_first_build_cache(&self, scope: TreeScope, condition: &UICondition, cache_request: &UICacheRequest) -> Result<UIElement> {
        let element = unsafe {
            self.element.FindFirstBuildCache(scope.into(), condition, cache_request)?
        };
        Ok(element.into())
    }

    /// Returns all UI Automation elements that satisfy the specified condition.
    pub fn find_all(&self, scope: TreeScope, condition: &UICondition) -> Result<Vec<UIElement>> {
        let elements = unsafe {
            self.element.FindAll(scope.into(), condition.as_ref())?
        };
        Self::to_elements(elements)
    }

    /// Returns all UI Automation elements that satisfy the specified condition, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    pub fn find_all_build_cache(&self, scope: TreeScope, condition: &UICondition, cache_request: &UICacheRequest) -> Result<Vec<UIElement>> {
        let elements = unsafe {
            self.element.FindAllBuildCache(scope.into(), condition, cache_request)?
        };
        Self::to_elements(elements)
    }

    /// Retrieves from the cache the parent of this UI Automation element.
    pub fn get_cached_parent(&self) -> Result<UIElement> {
        let parent = unsafe {
            self.element.GetCachedParent()?
        };
        Ok(parent.into())
    }

    /// Retrieves the cached child elements of this UI Automation element.
    pub fn get_cached_children(&self) -> Result<Vec<UIElement>> {
        let elements = unsafe {
            self.element.GetCachedChildren()?
        };
        Self::to_elements(elements)
    }

    /// Receives the runtime ID as a vec of integers.
    pub fn get_runtime_id(&self) -> Result<Vec<i32>> {
        let id = unsafe {
            self.element.GetRuntimeId()?
        };

        let arr = SafeArray::new(id, false);
        arr.try_into()
    }

    /// Retrieves the name of the element.
    pub fn get_name(&self) -> Result<String> {
        let name = unsafe {
            self.element.CurrentName()?
        };

        Ok(name.to_string())
    }

    /// Retrieves the cached name of the element.
    pub fn get_cached_name(&self) -> Result<String> {
        let name = unsafe {
            self.element.CachedName()?
        };
        Ok(name.to_string())
    }

    /// Retrieves the Microsoft UI Automation identifier of the element.
    pub fn get_automation_id(&self) -> Result<String> {
        let automation_id = unsafe {
            self.element.CurrentAutomationId()?
        };

        Ok(automation_id.to_string())
    }

    /// Retrieves the cached Microsoft UI Automation identifier of the element.
    pub fn get_cached_automation_id(&self) -> Result<String> {
        let automation_id = unsafe {
            self.element.CachedAutomationId()?
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

    /// Retrieves the cached ID of the process that hosts the element.
    pub fn get_cached_process_id(&self) -> Result<i32> {
        let id = unsafe {
            self.element.CachedProcessId()?
        };
        Ok(id)
    }

    /// Retrieves the class name of the element.
   pub fn get_classname(&self) -> Result<String> {
        let classname = unsafe {
            self.element.CurrentClassName()?
        };

        Ok(classname.to_string())
    }

    /// Retrieves the cached class name of the element.
    pub fn get_cached_classname(&self) -> Result<String> {
        let classname = unsafe {
            self.element.CachedClassName()?
        };
        Ok(classname.to_string())
    }

    /// Retrieves the control type of the element.
    pub fn get_control_type(&self) -> Result<ControlType> {
        let control_type = unsafe {
            self.element.CurrentControlType()?
        };

        Ok(ControlType::from(control_type))
    }

    /// Retrieves the cached control type of the element.
    pub fn get_cached_control_type(&self) -> Result<ControlType> {
        let control_type = unsafe {
            self.element.CachedControlType()?
        };
        Ok(control_type.into())
    }

    /// Retrieves a localized description of the control type of the element.
    pub fn get_localized_control_type(&self) -> Result<String> {
        let control_type = unsafe {
            self.element.CurrentLocalizedControlType()?
        };

        Ok(control_type.to_string())
    }

    /// Retrieves cached localized description of the control type of the element.
    pub fn get_cached_localized_control_type(&self) -> Result<String> {
        let control_type = unsafe {
            self.element.CachedLocalizedControlType()?
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

    /// Retrieves the cached accelerator key for the element.
    pub fn get_cached_accelerator_key(&self) -> Result<String> {
        let accelerator_key = unsafe {
            self.element.CachedAcceleratorKey()?
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

    /// Retrieves the cached access key character for the element.
    pub fn get_cached_access_key(&self) -> Result<String> {
        let access_key = unsafe {
            self.element.CachedAccessKey()?
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

    /// A cached value that indicates whether the element has keyboard focus.
    pub fn has_cached_keyboard_focus(&self) -> Result<bool> {
        let has_focus = unsafe {
            self.element.CachedHasKeyboardFocus()?
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

    /// A cached value that indicates whether the element can accept keyboard focus.
    pub fn is_cached_keyboard_focusable(&self) -> Result<bool> {
        let focusable = unsafe {
            self.element.CachedIsKeyboardFocusable()?
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

    /// A cached value that indicates whether the element is enabled.
    pub fn is_cached_enabled(&self) -> Result<bool> {
        let enabled = unsafe {
            self.element.CachedIsEnabled()?
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

    /// Retrieves the cached help text for the element.
    pub fn get_cached_help_text(&self) -> Result<String> {
        let text = unsafe {
            self.element.CachedHelpText()?
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

    /// Retrieves the cached culture identifier for the element.
    pub fn get_cached_culture(&self) -> Result<i32> {
        let culture = unsafe {
            self.element.CachedCulture()?
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

    /// A cached value that indicates whether the element is a control element.
    pub fn is_cached_control_element(&self) -> Result<bool> {
        let is_control = unsafe {
            self.element.CachedIsControlElement()?
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

    /// A cached value that indicates whether the element is a content element.
    pub fn is_cached_content_element(&self) -> Result<bool> {
        let is_content = unsafe {
            self.element.CachedIsContentElement()?
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

    /// A cached value that indicates whether the element contains a disguised password.
    pub fn is_cached_password(&self) -> Result<bool> {
        let is_password = unsafe {
            self.element.CachedIsPassword()?
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

    /// Retrieves the cached window handle of the element.
    pub fn get_cached_native_window_handle(&self) -> Result<Handle> {
        let handle = unsafe {
            self.element.CachedNativeWindowHandle()?
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

    /// Retrieves a cached description of the type of UI item represented by the element.
    pub fn get_cached_item_type(&self) -> Result<String> {
        let item_type = unsafe {
            self.element.CachedItemType()?
        };

        Ok(item_type.to_string())
    }

    /// Indicates whether the element is off-screen.
    pub fn is_offscreen(&self) -> Result<bool> {
        let off_screen = unsafe {
            self.element.CurrentIsOffscreen()?
        };

        Ok(off_screen.as_bool())
    }

    /// A cached value that indicates whether the element is off-screen.
    pub fn is_cached_offscreen(&self) -> Result<bool> {
        let off_screen = unsafe {
            self.element.CachedIsOffscreen()?
        };

        Ok(off_screen.as_bool())
    }

    /// Retrieves a value that indicates the orientation of the element.
    pub fn get_orientation(&self) -> Result<OrientationType> {
        let orientation = unsafe {
            self.element.CurrentOrientation()?
        };

        Ok(orientation.into())
    }

    /// Retrieves a cached value that indicates the orientation of the element.
    pub fn get_cached_orientation(&self) -> Result<OrientationType> {
        let orientation = unsafe {
            self.element.CachedOrientation()?
        };

        Ok(orientation.into())
    }

    /// Retrieves the name of the underlying UI framework.
    pub fn get_framework_id(&self) -> Result<String> {
        let id = unsafe {
            self.element.CurrentFrameworkId()?
        };

        Ok(id.to_string())
    }

    /// Retrieves the cached name of the underlying UI framework.
    pub fn get_cached_framework_id(&self) -> Result<String> {
        let id = unsafe {
            self.element.CachedFrameworkId()?
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

    /// A cached value that indicates whether the element is required to be filled out on a form.
    pub fn is_cached_required_for_form(&self) -> Result<bool> {
        let required = unsafe {
            self.element.CachedIsRequiredForForm()?
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

    /// A cached value that indicates whether the element contains valid data for a form.
    pub fn is_cached_data_valid_for_form(&self) -> Result<bool> {
        let valid = unsafe {
            self.element.CachedIsDataValidForForm()?
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

    /// Retrieves the cached description of the status of an item in an element.
    pub fn get_cached_item_status(&self) -> Result<String> {
        let status = unsafe {
            self.element.CachedItemStatus()?
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

    /// Retrieves the cached coordinates of the rectangle that completely encloses the element.
    pub fn get_cached_bounding_rectangle(&self) -> Result<Rect> {
        let rect = unsafe {
            self.element.CachedBoundingRectangle()?
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

    /// Retrieves the cached element that contains the text label for this element.
    pub fn get_cached_labeled_by(&self) -> Result<UIElement> {
        let labeled_by = unsafe {
            self.element.CachedLabeledBy()?
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

    /// Retrieves a cached array of elements for which this element serves as the controller.
    pub fn get_cached_controller_for(&self) -> Result<Vec<UIElement>> {
        let elements = unsafe {
            self.element.CachedControllerFor()?
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

    /// Retrieves a cached array of elements that describe this element.
    pub fn get_cached_described_by(&self) -> Result<Vec<UIElement>> {
        let elements = unsafe {
            self.element.CachedDescribedBy()?
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

    /// Retrieves a cached array of elements that indicates the reading order after the current element.
    pub fn get_cached_flows_to(&self) -> Result<Vec<UIElement>> {
        let elements = unsafe {
            self.element.CachedFlowsTo()?
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

    /// Retrieves a cached description of the provider for this element.
    pub fn get_cached_provider_description(&self) -> Result<String> {
        let descr = unsafe {
            self.element.CachedProviderDescription()?
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

    /// Try to set focus, return `true` if focus successfully.
    pub fn try_focus(&self) -> bool {
        self.set_focus().is_ok()
    }

    /// Retrieves the control pattern interface of the specified pattern `<T>` from this UI Automation element.
    pub fn get_pattern<T: UIPattern + TryFrom<IUnknown, Error = Error>>(&self) -> Result<T> {
        let pattern = unsafe {
            self.element.GetCurrentPattern(T::TYPE.into())?
        };

        T::try_from(pattern)
    }

    /// Retrieves the cached control pattern interface of the specified pattern `<T>` from this UI Automation element.
    pub fn get_cached_pattern<T: UIPattern + TryFrom<IUnknown, Error = Error>>(&self) -> Result<T> {
        let pattern = unsafe {
            self.element.GetCachedPattern(T::TYPE.into())?
        };

        T::try_from(pattern)
    }

    /// Retrieves a point on the element that can be clicked.
    pub fn get_clickable_point(&self) -> Result<Option<Point>> {
        let mut point = Point::default();
        let got = unsafe {
            self.element.GetClickablePoint(point.as_mut())?
        };

        Ok(if got.as_bool() {
            Some(point)
        } else {
            None
        })
    }

    /// Retrieves the current value of a property for this UI Automation element.
    pub fn get_property_value(&self, property: UIProperty) -> Result<Variant> {
        let value = unsafe {
            self.element.GetCurrentPropertyValue(property.into())?
        };

        Ok(value.into())
    }

    /// Retrieves the cached value of a property for this UI Automation element.
    pub fn get_cached_property_value(&self, property: UIProperty) -> Result<Variant> {
        let value = unsafe {
            self.element.GetCachedPropertyValue(property.into())?
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

    pub fn get_heading_level(&self) -> Result<HeadingLevel> {
        let element8: IUIAutomationElement8 = self.element.cast()?;
        let heading_level = unsafe {
            element8.CurrentHeadingLevel()?
        };
        Ok(heading_level.into())
    }

    pub fn get_cached_heading_level(&self) -> Result<HeadingLevel> {
        let element8: IUIAutomationElement8 = self.element.cast()?;
        let heading_level = unsafe {
            element8.CachedHeadingLevel()?
        };
        Ok(heading_level.into())
    }

    pub fn is_dialog(&self) -> Result<bool> {
        let element9: IUIAutomationElement9 = self.element.cast()?;
        let got = unsafe {
            element9.CurrentIsDialog()?
        };
        Ok(got.as_bool())
    }

    pub fn is_cached_dialog(&self) -> Result<bool> {
        let element9: IUIAutomationElement9 = self.element.cast()?;
        let got = unsafe {
            element9.CachedIsDialog()?
        };
        Ok(got.as_bool())
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

    /// Simulates typing `keys` on keyboard.
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

    /// Simulates holding `holdkeys` on keyboard, then sending `keys`.
    ///
    /// The key format is the same with `send_keys()`.
    ///
    /// # Examples
    ///
    /// ```
    /// use uiautomation::core::UIAutomation;
    ///
    /// let automation = UIAutomation::new().unwrap();
    /// let root = automation.get_root_element().unwrap();
    /// root.hold_send_keys("{Win}", "P", 0).unwrap();
    /// ```
    pub fn hold_send_keys(&self, holdkeys: &str, keys: &str, interval: u64) -> Result<()> {
        self.set_focus()?;

        let mut kb = Keyboard::new().interval(interval);

        kb.begin_hold_keys(holdkeys)?;
        kb.send_keys(keys)?;
        kb.end_hold_keys()
    }

    /// Simulates mouse left click event on the element.
    pub fn click(&self) -> Result<()> {
        self.try_focus();

        let point = self.get_click_point()?;
        let mouse = Mouse::default();
        mouse.click(point)
    }

    /// Simulates mouse left click event with holdkeys on the element.
    ///
    /// The holdkey is quoted by `{}`, for example: `{Ctrl}`, `{Ctrl}{Shift}`.
    pub fn hold_click(&self, holdkeys: &str) -> Result<()> {
        let point = self.get_click_point()?;
        let mouse = Mouse::default().holdkeys(holdkeys);
        mouse.click(point)
    }

    /// Simulates mouse double click event on the element.
    pub fn double_click(&self) -> Result<()> {
        self.try_focus();

        let point = self.get_click_point()?;
        let mouse = Mouse::default();
        mouse.double_click(point)
    }

    /// Simulates mouse right click event on the element.
    pub fn right_click(&self) -> Result<()> {
        self.try_focus();

        let point = self.get_click_point()?;
        let mouse = Mouse::default();
        mouse.right_click(point)
    }

    fn get_click_point(&self) -> Result<Point> {
        if let Ok(Some(point)) = self.get_clickable_point() {
            Ok(point)
        } else {
            let rect = self.get_bounding_rectangle()?;
            let point = Point::new((rect.get_left() + rect.get_right()) / 2, (rect.get_top() + rect.get_bottom()) / 2);
            Ok(point)
        }
    }
}

impl From<IUIAutomationElement> for UIElement {
    fn from(element: IUIAutomationElement) -> Self {
        UIElement {
            element
        }
    }
}

impl From<&IUIAutomationElement> for UIElement {
    fn from(value: &IUIAutomationElement) -> Self {
        value.clone().into()
    }
}

impl Into<IUIAutomationElement> for UIElement {
    fn into(self) -> IUIAutomationElement {
        self.element
    }
}

impl Param<IUIAutomationElement> for UIElement {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationElement> {
        windows::core::ParamValue::Owned(self.element)
    }
}

impl Param<IUIAutomationElement> for &UIElement {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationElement> {
        windows::core::ParamValue::Borrowed(self.element.as_raw())
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

/// Exposes properties and methods of a cache request.
/// Client applications use this interface to specify the properties and control patterns to be cached when a Microsoft UI Automation element is obtained.
#[derive(Debug, Clone)]
pub struct UICacheRequest {
    request: IUIAutomationCacheRequest
}

impl UICacheRequest {
    /// Adds a control pattern to the cache request.
    pub fn add_pattern(&self, pattern: UIPatternType) -> Result<()> {
        unsafe {
            self.request.AddPattern(pattern.into())?
        };
        Ok(())
    }

    /// Adds a property to the cache request.
    pub fn add_property(&self, property: UIProperty) -> Result<()> {
        unsafe {
            self.request.AddProperty(property.into())?
        };
        Ok(())
    }

    /// Retrieves whether returned elements contain full references to the underlying UI, or only cached information.
    pub fn get_element_mode(&self) -> Result<ElementMode> {
        let mode = unsafe {
            self.request.AutomationElementMode()?
        };
        Ok(mode.into())
    }

    /// Sets whether returned elements contain full references to the underlying UI, or only cached information.
    pub fn set_element_mode(&self, mode: ElementMode) -> Result<()> {
        unsafe {
            self.request.SetAutomationElementMode(mode.into())?
        };
        Ok(())
    }

    /// Retrieves the view of the UI Automation element tree that is used when caching.
    pub fn get_tree_filter(&self) -> Result<UICondition> {
        let condition = unsafe {
            self.request.TreeFilter()?
        };
        Ok(condition.into())
    }

    /// Sets the view of the UI Automation element tree that is used when caching.
    pub fn set_tree_filter(&self, filter: UICondition) -> Result<()> {
        unsafe {
            self.request.SetTreeFilter(filter)?
        };
        Ok(())
    }

    /// Retrieves the scope of caching.
    pub fn get_tree_scope(&self) -> Result<TreeScope> {
        let scope = unsafe {
            self.request.TreeScope()?
        };
        Ok(scope.into())
    }

    /// Sets the scope of caching.
    pub fn set_tree_scope(&self, scope: TreeScope) -> Result<()> {
        unsafe {
            self.request.SetTreeScope(scope.into())?
        };
        Ok(())
    }
}

impl From<IUIAutomationCacheRequest> for UICacheRequest {
    fn from(value: IUIAutomationCacheRequest) -> Self {
        Self {
            request: value
        }
    }
}

impl Into<IUIAutomationCacheRequest> for UICacheRequest {
    fn into(self) -> IUIAutomationCacheRequest {
        self.request
    }
}

impl AsRef<IUIAutomationCacheRequest> for UICacheRequest {
    fn as_ref(&self) -> &IUIAutomationCacheRequest {
        &self.request
    }
}

impl Param<IUIAutomationCacheRequest> for UICacheRequest {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationCacheRequest> {
        windows::core::ParamValue::Owned(self.request)
    }
}

impl Param<IUIAutomationCacheRequest> for &UICacheRequest {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationCacheRequest> {
        windows::core::ParamValue::Borrowed(self.request.as_raw())
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

    /// Retrieves the parent element of the specified UI Automation element, and caches properties and control patterns.
    pub fn get_parent_build_cache(&self, element: &UIElement, cache_request: &UICacheRequest) -> Result<UIElement> {
        let parent = unsafe {
            self.tree_walker.GetParentElementBuildCache(element, cache_request)?
        };
        Ok(parent.into())
    }

    /// Retrieves the first child element of the specified UI Automation element.
    pub fn get_first_child(&self, element: &UIElement) -> Result<UIElement> {
        let child: IUIAutomationElement;
        unsafe {
            child = self.tree_walker.GetFirstChildElement(&element.element)?;
        }

        Ok(UIElement::from(child))
    }

    /// Retrieves the first child element of the specified UI Automation element, and caches properties and control patterns.
    pub fn get_first_child_build_cache(&self, element: &UIElement, cache_request: &UICacheRequest) -> Result<UIElement> {
        let child = unsafe {
            self.tree_walker.GetFirstChildElementBuildCache(element, cache_request)?
        };
        Ok(child.into())
    }

    /// Retrieves the last child element of the specified UI Automation element.
    pub fn get_last_child(&self, element: &UIElement) -> Result<UIElement> {
        let child: IUIAutomationElement;
        unsafe {
            child = self.tree_walker.GetLastChildElement(&element.element)?;
        }

        Ok(UIElement::from(child))
    }

    /// Retrieves the last child element of the specified UI Automation element, and caches properties and control patterns.
    pub fn get_last_child_build_cache(&self, element: &UIElement, cache_request: &UICacheRequest) -> Result<UIElement> {
        let child = unsafe {
            self.tree_walker.GetLastChildElementBuildCache(element, cache_request)?
        };
        Ok(child.into())
    }

    /// Retrieves the next sibling element of the specified UI Automation element.
    pub fn get_next_sibling(&self, element: &UIElement) -> Result<UIElement> {
        let sibling: IUIAutomationElement;
        unsafe {
            sibling = self.tree_walker.GetNextSiblingElement(&element.element)?;
        }

        Ok(UIElement::from(sibling))
    }

    /// Retrieves the next sibling element of the specified UI Automation element, and caches properties and control patterns.
    pub fn get_next_sibling_build_cache(&self, element: &UIElement, cache_request: &UICacheRequest) -> Result<UIElement> {
        let sibling = unsafe {
            self.tree_walker.GetNextSiblingElementBuildCache(element, cache_request)?
        };
        Ok(sibling.into())
    }

    /// Retrieves the previous sibling element of the specified UI Automation element.
    pub fn get_previous_sibling(&self, element: &UIElement) -> Result<UIElement> {
        let sibling: IUIAutomationElement;
        unsafe {
            sibling = self.tree_walker.GetPreviousSiblingElement(&element.element)?;
        }

        Ok(UIElement::from(sibling))
    }

    /// Retrieves the previous sibling element of the specified UI Automation element, and caches properties and control patterns.
    pub fn get_previous_sibling_build_cache(&self, element: &UIElement, cache_request: &UICacheRequest) -> Result<UIElement> {
        let sibling = unsafe {
            self.tree_walker.GetPreviousSiblingElementBuildCache(element, cache_request)?
        };
        Ok(sibling.into())
    }

    /// Retrieves the ancestor element nearest to the specified Microsoft UI Automation element in the tree view.
    pub fn normalize(&self, element: &UIElement) -> Result<UIElement> {
        let result = unsafe {
            self.tree_walker.NormalizeElement(element.as_ref())?
        };
        Ok(result.into())
    }

    /// Retrieves the ancestor element nearest to the specified Microsoft UI Automation element in the tree view,
    /// prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    pub fn normalize_build_cache(&self, element: &UIElement, cache_request: &UICacheRequest) -> Result<UIElement> {
        let ret = unsafe {
            self.tree_walker.NormalizeElementBuildCache(element, cache_request)?
        };
        Ok(ret.into())
    }

    /// Retrieves the condition that defines the view of the UI Automation tree.
    pub fn get_condition(&self) -> Result<UICondition> {
        let condition = unsafe {
            self.tree_walker.Condition()?
        };
        Ok(condition.into())
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

/// Defines the uielement mode when matcher is searching for.
#[derive(Debug)]
pub enum UIMatcherMode {
    /// Searches all element.
    Raw,
    /// Searches control element only.
    Control,
    /// Searches content element only.
    Content
}

/// Defines filter conditions to match specific UI Element.
///
/// `UIMatcher` can find first element or find all elements.
pub struct UIMatcher {
    automation: UIAutomation,
    mode: UIMatcherMode,
    depth: u32,
    from: Option<UIElement>,
    // condition: Option<Box<dyn Condition>>,
    filters: Vec<Box<dyn MatcherFilter>>,
    timeout: u64,
    interval: u64,
    debug: bool
}

impl UIMatcher {
    /// Creates a matcher with `automation`.
    pub fn new(automation: UIAutomation) -> Self {
        UIMatcher {
            automation,
            mode: UIMatcherMode::Control,
            depth: 7,
            from: None,
            filters: Vec::new(),
            timeout: 3000,
            interval: 100,
            debug: false
        }
    }

    /// Sets the searching mode. `UIMatcherMode::Control` is default mode.
    pub fn mode(mut self, search_mode: UIMatcherMode) -> Self {
        self.mode = search_mode;
        self
    }

    /// Sets the root element of the UIAutomation tree whitch should be searched from.
    ///
    /// The root element is desktop by default.
    pub fn from(mut self, element: UIElement) -> Self {
        self.from = Some(element);
        self
    }

    /// Sets the root element of the UIAutomation tree whitch should be searched from. The `element` is cloned internally.
    ///
    /// The root element is desktop by default.
    pub fn from_ref(mut self, element: &UIElement) -> Self {
        self.from = Some(element.clone());
        self
    }

    /// Sets the depth of the search path. The default depth is `7`.
    pub fn depth(mut self, depth: u32) -> Self {
        self.depth = depth;
        self
    }

    /// Sets the the time in millionseconds for matching element. The default timeout is 3000 millionseconds(3 seconds).
    ///
    /// The `UIMatcher` will not retry to find when you set `timeout` to `0`.
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
     pub fn filter(mut self, filter: Box<dyn MatcherFilter>) -> Self {
        self.filters.push(filter);
        self
    }

    /// Appends a filter function which is used as `and` logic.
    ///
    /// # Examples:
    ///
    /// ```
    /// use uiautomation::core::UIAutomation;
    /// use uiautomation::core::UIElement;
    ///
    /// let automation = UIAutomation::new().unwrap();
    /// let matcher = automation.create_matcher().filter_fn(Box::new(|e: &UIElement| {
    ///     let framework_id = e.get_framework_id()?;
    ///     let class_name = e.get_classname()?;
    ///
    ///     Ok("Win32" == framework_id && class_name.starts_with("Shell"))
    /// })).timeout(0);
    /// let element = matcher.find_first();
    /// assert!(element.is_ok());
    /// ```
    pub fn filter_fn<F>(mut self, f: Box<F>) -> Self where F: Fn(&UIElement) -> Result<bool> + 'static {
        let filter = FnFilter {
            filter: f
        };
        self.filters.push(Box::new(filter));
        self
    }

    /// Append a filter whitch match specific casesensitive name.
    pub fn name<S: Into<String>>(self, name: S) -> Self {
        let condition = NameFilter {
            value: name.into(),
            casesensitive: true,
            partial: false
        };

        self.filter(Box::new(condition))
    }

    /// Append a filter whitch name contains specific text (ignore casesensitive).
    pub fn contains_name<S: Into<String>>(self, name: S) -> Self {
        let condition = NameFilter {
            value: name.into(),
            casesensitive: false,
            partial: true
        };
        self.filter(Box::new(condition))
    }

    /// Append a filter whitch matches specific name (ignore casesensitive).
    pub fn match_name<S: Into<String>>(self, name: S) -> Self {
        let condition = NameFilter {
            value: name.into(),
            casesensitive: false,
            partial: false
        };
        self.filter(Box::new(condition))
    }

    /// Filters by classname.
    pub fn classname<S: Into<String>>(self, classname: S) -> Self {
        let condition = ClassNameFilter {
            classname: classname.into()
        };
        self.filter(Box::new(condition))
    }

    /// Filters by control type.
    pub fn control_type(self, control_type: ControlType) -> Self {
        let condition = ControlTypeFilter {
            control_type
        };
        self.filter(Box::new(condition))
    }

    /// Clears all filters.
    pub fn reset(mut self) -> Self {
        // self.condition = None;
        self.filters.clear();
        self
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
                    #[cfg(feature = "log")]
                    log::debug!("Try to match element...");
                    #[cfg(not(feature = "log"))]
                    println!("Try to match element...")
                }

            let (root, walker) = self.prepare()?;
            self.search(&walker, &root, &mut elements, 1, first_only)?;

            if !elements.is_empty() || self.timeout <= 0 {
                break;
            }

            let now = Local::now().timestamp_millis();
            if now - start >= self.timeout as i64 {
                return Err(Error::new(ERR_TIMEOUT, "find element time out"));
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
        let walker = match self.mode {
            UIMatcherMode::Raw => self.automation.create_tree_walker()?,
            UIMatcherMode::Control => self.automation.filter_tree_walker(self.automation.get_control_view_condition()?)?,
            UIMatcherMode::Content => self.automation.filter_tree_walker(self.automation.get_content_view_condition()?)?,
        };

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
        if let Some(ref root) = self.from {
            if self.automation.compare_elements(root, element)? {
                return Ok(false);
            }
        }

        // let ret = if let Some(ref condition) = self.condition {
        //     condition.judge(element)?
        // } else {
        //     true
        // };

        let mut ret = true;
        let mut failed_filter = 0;
        for (idx, condition) in self.filters.iter().enumerate() {
            ret = condition.judge(element)?;
            if !ret {
                failed_filter = idx;
                break;
            }
        }

        if self.debug {
            #[cfg(feature = "log") ]
            log::debug!("{:?} -> {} in filter {}", element, ret, failed_filter);
            #[cfg(not(feature = "log"))]
            println!("{:?} -> {} in filter {}", element, ret, failed_filter);
        }

        Ok(ret)
    }
}

impl Debug for UIMatcher {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("UIMatcher")
            .field("automation", &self.automation)
            .field("mode", &self.mode)
            .field("depth", &self.depth)
            .field("from", &self.from)
            .field("filters", &format!("({} filers)", self.filters.len()))
            .field("timeout", &self.timeout)
            .field("interval", &self.interval)
            .field("debug", &self.debug)
        .finish()
    }
}

/// This is the trait for conditions used in filtering when searching for elements in the UI Automation tree.
pub trait IUICondition<T: Interface>: Sized + From<T> + Into<T> + AsRef<T> {
}

/// This is the primary interface for conditions used in filtering when searching for elements in the UI Automation tree.
#[derive(Debug, Clone)]
pub struct UICondition(IUIAutomationCondition);

impl UICondition {
    /// Determines whether the condition is `bool` condition.
    pub fn is_bool_condition(&self) -> bool {
        let bool_cond: windows::core::Result<IUIAutomationBoolCondition> = self.0.cast();
        bool_cond.is_ok()
    }

    /// Determines whether the condition is `and` condition.
    pub fn is_and_condition(&self) -> bool {
        let and_cond: windows::core::Result<IUIAutomationAndCondition> = self.0.cast();
        and_cond.is_ok()
    }

    /// Determines whether the condition is `and` condition.
    pub fn is_or_condition(&self) -> bool {
        let or_cond: windows::core::Result<IUIAutomationOrCondition> = self.0.cast();
        or_cond.is_ok()
    }

    /// Determines whether the condition is `not` condition.
    pub fn is_not_condition(&self) -> bool {
        let not_cond: windows::core::Result<IUIAutomationNotCondition> = self.0.cast();
        not_cond.is_ok()
    }

    /// Determines whether the condition is `property` condition.
    pub fn is_property_condition(&self) -> bool {
        let prop_cond: windows::core::Result<IUIAutomationPropertyCondition> = self.0.cast();
        prop_cond.is_ok()
    }
}

impl IUICondition<IUIAutomationCondition> for UICondition {
}

impl From<IUIAutomationCondition> for UICondition {
    fn from(condition: IUIAutomationCondition) -> Self {
        Self(condition)
    }
}

impl Into<IUIAutomationCondition> for UICondition {
    fn into(self) -> IUIAutomationCondition {
        self.0
    }
}

impl AsRef<IUIAutomationCondition> for UICondition {
    fn as_ref(&self) -> &IUIAutomationCondition {
        &self.0
    }
}

impl Param<IUIAutomationCondition> for UICondition {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationCondition> {
        windows::core::ParamValue::Owned(self.0)
    }
}

impl Param<IUIAutomationCondition> for &UICondition {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationCondition> {
        windows::core::ParamValue::Borrowed(self.0.as_raw())
    }
}

/// Represents a condition that can be either TRUE (selects all elements) or FALSE (selects no elements).
#[derive(Debug, Clone)]
pub struct UIBoolCondition(IUIAutomationBoolCondition);

impl UIBoolCondition {
    pub fn get_bool_value(&self) -> Result<bool> {
        let val = unsafe {
            self.0.BooleanValue()?
        };
        Ok(val.as_bool())
    }
}

impl IUICondition<IUIAutomationBoolCondition> for UIBoolCondition {
}

impl From<IUIAutomationBoolCondition> for UIBoolCondition {
    fn from(condition: IUIAutomationBoolCondition) -> Self {
        Self(condition)
    }
}

impl Into<IUIAutomationBoolCondition> for UIBoolCondition {
    fn into(self) -> IUIAutomationBoolCondition {
        self.0
    }
}

impl AsRef<IUIAutomationBoolCondition> for UIBoolCondition {
    fn as_ref(&self) -> &IUIAutomationBoolCondition {
        &self.0
    }
}

impl TryFrom<IUIAutomationCondition> for UIBoolCondition {
    type Error = Error;

    fn try_from(condition: IUIAutomationCondition) -> Result<Self> {
        let bool_cond: IUIAutomationBoolCondition = condition.cast()?;
        Ok(bool_cond.into())
    }
}

impl Into<IUIAutomationCondition> for UIBoolCondition {
    fn into(self) -> IUIAutomationCondition {
        self.0.cast().unwrap()
    }
}

impl TryFrom<UICondition> for UIBoolCondition {
    type Error = Error;

    fn try_from(condition: UICondition) -> Result<Self> {
        condition.0.try_into()
    }
}

impl Into<UICondition> for UIBoolCondition {
    fn into(self) -> UICondition {
        let condition: IUIAutomationCondition = self.0.cast().unwrap();
        condition.into()
    }
}

/// Represents a condition that is the negative of another condition.
#[derive(Debug, Clone)]
pub struct UINotCondition(IUIAutomationNotCondition);

impl UINotCondition {
    /// Retrieves the condition of which this condition is the negative.
    pub fn get_child(&self) -> Result<UICondition> {
        let child = unsafe {
            self.0.GetChild()?
        };
        Ok(child.into())
    }
}

impl IUICondition<IUIAutomationNotCondition> for UINotCondition {
}

impl From<IUIAutomationNotCondition> for UINotCondition {
    fn from(condition: IUIAutomationNotCondition) -> Self {
        Self(condition)
    }
}

impl Into<IUIAutomationNotCondition> for UINotCondition {
    fn into(self) -> IUIAutomationNotCondition {
        self.0
    }
}

impl AsRef<IUIAutomationNotCondition> for UINotCondition {
    fn as_ref(&self) -> &IUIAutomationNotCondition {
        &self.0
    }
}

impl TryFrom<IUIAutomationCondition> for UINotCondition {
    type Error = Error;

    fn try_from(condition: IUIAutomationCondition) -> Result<Self> {
        let not_cond: IUIAutomationNotCondition = condition.cast()?;
        Ok(not_cond.into())
    }
}

impl Into<IUIAutomationCondition> for UINotCondition {
    fn into(self) -> IUIAutomationCondition {
        self.0.cast().unwrap()
    }
}

impl TryFrom<UICondition> for UINotCondition {
    type Error = Error;

    fn try_from(condition: UICondition) -> Result<Self> {
        condition.0.try_into()
    }
}

impl Into<UICondition> for UINotCondition {
    fn into(self) -> UICondition {
        let condition: IUIAutomationCondition = self.0.cast().unwrap();
        condition.into()
    }
}

/// Exposes properties and methods that Microsoft UI Automation client applications can use to retrieve information about an AND-based property condition.
#[derive(Debug, Clone)]
pub struct UIAndCondition(IUIAutomationAndCondition);

impl UIAndCondition {
    /// Retrieves the number of conditions that make up this "and" condition.
    pub fn get_children_count(&self) -> Result<i32> {
        let count = unsafe {
            self.0.ChildCount()?
        };
        Ok(count)
    }

    /// Retrieves the conditions that make up this "and" condition.
    pub fn get_children(&self) -> Result<Vec<UICondition>> {
        let arr = unsafe {
            self.0.GetChildren()?
        };
        let arr = SafeArray::from(arr);
        let children: Vec<IUIAutomationCondition> = arr.into_interface_vector()?;
        let conditions: Vec<UICondition> = children.into_iter().map(|c| c.into()).collect();
        Ok(conditions)
    }
}

impl IUICondition<IUIAutomationAndCondition> for UIAndCondition {
}

impl From<IUIAutomationAndCondition> for UIAndCondition {
    fn from(condition: IUIAutomationAndCondition) -> Self {
        Self(condition)
    }
}

impl Into<IUIAutomationAndCondition> for UIAndCondition {
    fn into(self) -> IUIAutomationAndCondition {
        self.0
    }
}

impl AsRef<IUIAutomationAndCondition> for UIAndCondition {
    fn as_ref(&self) -> &IUIAutomationAndCondition {
        &self.0
    }
}

impl TryFrom<IUIAutomationCondition> for UIAndCondition {
    type Error = Error;

    fn try_from(condition: IUIAutomationCondition) -> Result<Self> {
        let and_cond: IUIAutomationAndCondition = condition.cast()?;
        Ok(and_cond.into())
    }
}

impl Into<IUIAutomationCondition> for UIAndCondition {
    fn into(self) -> IUIAutomationCondition {
        self.0.cast().unwrap()
    }
}

impl TryFrom<UICondition> for UIAndCondition {
    type Error = Error;

    fn try_from(condition: UICondition) -> Result<Self> {
        condition.0.try_into()
    }
}

impl Into<UICondition> for UIAndCondition {
    fn into(self) -> UICondition {
        let condition: IUIAutomationCondition = self.0.cast().unwrap();
        condition.into()
    }
}

#[derive(Debug, Clone)]
pub struct UIOrCondition(IUIAutomationOrCondition);

impl UIOrCondition {
    /// Retrieves the number of conditions that make up this "or" condition.
    pub fn get_children_count(&self) -> Result<i32> {
        let count = unsafe {
            self.0.ChildCount()?
        };
        Ok(count)
    }

    /// Retrieves the conditions that make up this "or" condition.
    pub fn get_children(&self) -> Result<Vec<UICondition>> {
        let arr = unsafe {
            self.0.GetChildren()?
        };
        let arr = SafeArray::from(arr);
        let children: Vec<IUIAutomationCondition> = arr.into_interface_vector()?;
        let conditions: Vec<UICondition> = children.into_iter().map(|c| c.into()).collect();
        Ok(conditions)
    }
}

impl IUICondition<IUIAutomationOrCondition> for UIOrCondition {
}

impl From<IUIAutomationOrCondition> for UIOrCondition {
    fn from(condition: IUIAutomationOrCondition) -> Self {
        Self(condition)
    }
}

impl Into<IUIAutomationOrCondition> for UIOrCondition {
    fn into(self) -> IUIAutomationOrCondition {
        self.0
    }
}

impl AsRef<IUIAutomationOrCondition> for UIOrCondition {
    fn as_ref(&self) -> &IUIAutomationOrCondition {
        &self.0
    }
}

impl TryFrom<IUIAutomationCondition> for UIOrCondition {
    type Error = Error;

    fn try_from(condition: IUIAutomationCondition) -> Result<Self> {
        let or_cond: IUIAutomationOrCondition = condition.cast()?;
        Ok(or_cond.into())
    }
}

impl Into<IUIAutomationCondition> for UIOrCondition {
    fn into(self) -> IUIAutomationCondition {
        self.0.cast().unwrap()
    }
}

impl TryFrom<UICondition> for UIOrCondition {
    type Error = Error;

    fn try_from(condition: UICondition) -> Result<Self> {
        let or_cond: IUIAutomationOrCondition = condition.0.cast()?;
        Ok(or_cond.into())
    }
}

impl Into<UICondition> for UIOrCondition {
    fn into(self) -> UICondition {
        let condition: IUIAutomationCondition = self.into();
        condition.into()
    }
}

/// Represents a condition based on a property value that is used to find UI Automation elements.
#[derive(Debug, Clone)]
pub struct UIPropertyCondition(IUIAutomationPropertyCondition);

impl UIPropertyCondition {
    /// Retrieves the identifier of the property on which this condition is based.
    pub fn get_property(&self) -> Result<UIProperty> {
        let property_id = unsafe {
            self.0.PropertyId()?
        };
        Ok(property_id.into())
    }

    /// Retrieves the property value that must be matched for the condition to be true.
    pub fn get_property_value(&self) -> Result<Variant> {
        let value = unsafe {
            self.0.PropertyValue()?
        };
        Ok(value.into())
    }

    /// Retrieves a set of flags that specify how the condition is applied.
    pub fn get_property_condition_flags(&self) -> Result<PropertyConditionFlags> {
        Ok(unsafe {
            self.0.PropertyConditionFlags()?
        }.into())
    }
}

impl IUICondition<IUIAutomationPropertyCondition> for UIPropertyCondition {
}

impl From<IUIAutomationPropertyCondition> for UIPropertyCondition {
    fn from(condition: IUIAutomationPropertyCondition) -> Self {
        Self(condition)
    }
}

impl Into<IUIAutomationPropertyCondition> for UIPropertyCondition {
    fn into(self) -> IUIAutomationPropertyCondition {
        self.0
    }
}

impl AsRef<IUIAutomationPropertyCondition> for UIPropertyCondition {
    fn as_ref(&self) -> &IUIAutomationPropertyCondition {
        &self.0
    }
}

impl TryFrom<IUIAutomationCondition> for UIPropertyCondition {
    type Error = Error;

    fn try_from(condition: IUIAutomationCondition) -> Result<Self> {
        let prop_cond: IUIAutomationPropertyCondition = condition.cast()?;
        Ok(prop_cond.into())
    }
}

impl Into<IUIAutomationCondition> for UIPropertyCondition {
    fn into(self) -> IUIAutomationCondition {
        self.0.cast().unwrap()
    }
}

impl TryFrom<UICondition> for UIPropertyCondition {
    type Error = Error;

    fn try_from(condition: UICondition) -> Result<Self> {
        condition.0.try_into()
    }
}

impl Into<UICondition> for UIPropertyCondition {
    fn into(self) -> UICondition {
        let condition: IUIAutomationCondition = self.0.cast().unwrap();
        condition.into()
    }
}

#[cfg(test)]
mod tests {
    use std::thread::sleep;
    use std::time::Duration;

    use windows::Win32::UI::Accessibility::*;
    use windows::Win32::UI::WindowsAndMessaging::GetForegroundWindow;

    use crate::UIAutomation;
    use crate::UIElement;
    use crate::controls::ControlType;
    use crate::filters::MatcherFilter;
    use crate::types::TreeScope;

    fn print_element(element: &UIElement) {
        println!("Name: {}", element.get_name().unwrap());
        println!("ControlType: {:?}", element.get_control_type().unwrap());
        println!("LocalizedControlType: {}", element.get_localized_control_type().unwrap());
        println!("BoundingRectangle: {}", element.get_bounding_rectangle().unwrap());
        println!("IsEnabled: {}", element.is_enabled().unwrap());
        println!("IsOffscreen: {}", element.is_offscreen().unwrap());
        println!("IsKeyboardFocusable: {}", element.is_keyboard_focusable().unwrap());
        println!("HasKeyboardFocus: {}", element.has_keyboard_focus().unwrap());
        println!("ProcessId: {}", element.get_process_id().unwrap());
        println!("RuntimeId: {:?}", element.get_runtime_id().unwrap());
        println!("FrameworkId: {}", element.get_framework_id().unwrap());
        println!("ClassName: {}", element.get_classname().unwrap());
        println!("NativeWindowHandle: {}", element.get_native_window_handle().unwrap());
        println!("ProviderDescription: {}", element.get_provider_description().unwrap());
        println!("IsPassword: {}", element.is_password().unwrap());
        println!("AutomationId: {}", element.get_automation_id().unwrap());
        println!("HeadingLevel: {}", element.get_heading_level().unwrap());
        println!("IsDialog: {}", element.is_dialog().unwrap());
    }

    #[test]
    fn test_element_properties() {
        let automation = UIAutomation::new().unwrap();

        sleep(Duration::new(0, 500));

        let root = automation.get_root_element().unwrap();

        println!("---------------------");
        print_element(&root);
        println!("---------------------");
    }

    #[test]
    fn test_tree_walker() {
        let automation = UIAutomation::new().unwrap();
        let root = automation.get_root_element().unwrap();
        let walker = automation.create_tree_walker().unwrap();
        let child = walker.get_first_child(&root);
        assert!(child.is_ok());
    }

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
            let matcher = automation.create_matcher().control_type(ControlType::MenuItem).from_ref(&notepad).name("文件").depth(5).timeout(1000);
            if let Ok(menu_item) = matcher.find_first() {
                menu_item.click().unwrap();
            }
        }
    }

    #[test]
    fn test_element_find() {
        let automation = UIAutomation::new().unwrap();
        let root = automation.get_root_element().unwrap();
        let condition = automation.create_true_condition().unwrap();
        let child = root.find_first(TreeScope::Children, &condition).unwrap();
        println!("{}", child);
    }

    #[test]
    fn test_notepad() {
        let automation = UIAutomation::new().unwrap();
        let matcher = automation.create_matcher();
        if let Ok(window) = matcher.classname("Notepad").timeout(0).find_first() {
            println!("{}", window.get_name().unwrap());

            let menubar = automation.create_matcher() //.debug(true)
                .from(window.clone())
                .control_type(ControlType::Pane)
                .timeout(0)
                .find_first().unwrap();

            println!("{}, {}", menubar.get_framework_id().unwrap(), menubar.get_classname().unwrap());
        }
    }

    #[test]
    fn test_search_from() {
        let automation = UIAutomation::new().unwrap();
        if let Ok(window) = automation.create_matcher().classname("Notepad").find_first() {
            let nothing = automation.create_matcher().from(window.clone()).control_type(ControlType::Window).find_first();
            assert!(nothing.is_err());
        }
    }

    struct FrameworkIdFilter(String);

    impl MatcherFilter for FrameworkIdFilter {
        fn judge(&self, element: &crate::UIElement) -> crate::Result<bool> {
            let id = element.get_framework_id()?;
            Ok(id == self.0)
        }
    }

    #[test]
    fn test_custom_search() {
        let automation = UIAutomation::new().unwrap();
        let matcher = automation.create_matcher().timeout(0).filter(Box::new(FrameworkIdFilter("Win32".into()))).depth(2);
        let element = matcher.find_first();
        assert!(element.is_ok());
        println!("{}", element.unwrap());
    }

    #[test]
    fn test_find_no_wait() {
        let automation = UIAutomation::new().unwrap();
        let matcher = automation.create_matcher().timeout(0).name("You can find nothing!");
        let item = matcher.find_first();
        assert!(item.is_err());
    }

    #[test]
    fn test_automation_id() {
        let automation = UIAutomation::new().unwrap();
        if let Ok(notepad) = automation.create_matcher().timeout(0).classname("Notepad").find_first() {
            let title_bar = automation.create_matcher().from(notepad).timeout(0).control_type(ControlType::TitleBar).find_first().unwrap();
            let element: &IUIAutomationElement = title_bar.as_ref();
            let automation_id = unsafe {
                element.CurrentAutomationId().unwrap()
            };
            println!("{} -> {}", title_bar, automation_id);
        }
    }

    #[test]
    fn test_window_rect_prop() {
        let window = unsafe { GetForegroundWindow() };
        if window.is_invalid() {
            return;
        }

        let automation = UIAutomation::new().unwrap();
        let element = automation.element_from_handle(window.into()).unwrap();
        let rect = element.get_bounding_rectangle().unwrap();
        println!("Window Rect = {}", rect);

        let val = element.get_property_value(crate::types::UIProperty::BoundingRectangle).unwrap();
        println!("Window Rect Prop = {}", val.to_string());
        assert!(val.is_array());

        let arr = val.get_array().unwrap();
        let l: f64 = arr.get_element(0).unwrap();
        let t: f64 = arr.get_element(1).unwrap();
        let r: f64 = arr.get_element(2).unwrap();
        let b: f64 = arr.get_element(3).unwrap();
        println!("Window Rect Array = [{}, {}, {}, {}]", l, t, r, b);
    }

    #[test]
    fn test_create() {
        let _ = UIAutomation::new();

        let uiautomation = UIAutomation::new_direct();
        assert!(uiautomation.is_ok());
    }

    #[test]
    fn test_search_from_handle() {
        let auto = UIAutomation::new().unwrap();
        let _ = auto.element_from_handle(crate::types::Handle::from(0x2006C6));
    }
}
