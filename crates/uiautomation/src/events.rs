use uiautomation_derive::map_as;
use uiautomation_derive::EnumConvert;
use windows::Win32::UI::Accessibility::IUIAutomationElement;
use windows::Win32::UI::Accessibility::IUIAutomationEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationEventHandler_Impl;
use windows::Win32::UI::Accessibility::IUIAutomationFocusChangedEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationPropertyChangedEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationStructureChangedEventHandler;
use windows::Win32::UI::Accessibility::UIA_EVENT_ID;
use windows_core::implement;
use windows_core::Param;

use crate::types::UIProperty;
use crate::variants::SafeArray;
use crate::variants::Variant;
use crate::Result;
use crate::UIElement;

/// `UIEventType` is an enum wrapper for `windows::Win32::UI::Accessibility::UIA_EVENT_ID`.
/// 
/// Describes  the named constants used to identify Microsoft UI Automation events.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::UIA_EVENT_ID)]
#[allow(non_camel_case_types)]
pub enum UIEventType {
    /// Identifies the event that is raised when a tooltip is opened.
    ToolTipOpened = 20000i32,
    /// Identifies the event that is raised when a tooltip is closed.
    ToolTipClosed = 20001i32,
    /// Identifies the event that is raised when the UI Automation tree structure is changed.
    StructureChanged = 20002i32,
    /// Identifies the event that is raised when a menu is opened.
    MenuOpened = 20003i32,
    /// Identifies the event that is raised when the value of a property has changed.
    AutomationPropertyChanged = 20004i32,
    /// Identifies the event that is raised when the focus has changed from one element to another.
    AutomationFocusChanged = 20005i32,
    /// Identifies the event that is raised when asynchronous content is being loaded. This event is used mainly by providers to indicate that asynchronous content-loading events have occurred.
    AsyncContentLoaded = 20006i32,
    /// Identifies the event that is raised when a menu is closed.
    MenuClosed = 20007i32,
    /// Identifies the event that is raised when the layout of child items within a control has changed. This event is also used for Auto-suggest accessibility.
    LayoutInvalidated = 20008i32,
    /// Identifies the event that is raised when a control is invoked or activated.
    Invoke_Invoked = 20009i32,
    /// Identifies the event raised when an item is added to a collection of selected items.
    SelectionItem_ElementAddedToSelection = 20010i32,
    /// Identifies the event raised when an item is removed from a collection of selected items.
    SelectionItem_ElementRemovedFromSelection = 20011i32,
    /// Identifies the event that is raised when a call to the Select, AddToSelection, or RemoveFromSelection method results in a single item being selected.
    SelectionItem_ElementSelected = 20012i32,
    /// Identifies the event that is raised when a selection in a container has changed significantly.
    Selection_Invalidated = 20013i32,
    /// Identifies the event that is raised when the text selection is modified.
    Text_TextSelectionChanged = 20014i32,
    /// Identifies the event that is raised whenever textual content is modified.
    Text_TextChanged = 20015i32,
    /// Identifies the event that is raised when a window is opened.
    Window_WindowOpened = 20016i32,
    /// Identifies the event that is raised when a window is closed.
    Window_WindowClosed = 20017i32,
    /// Identifies the event that is raised when a menu mode is started.
    MenuModeStart = 20018i32,
    /// Identifies the event that is raised when a menu mode is ended.
    MenuModeEnd = 20019i32,
    /// Identifies the event that is raised when the specified mouse or keyboard input reaches the element for which the StartListening method was called.
    InputReachedTarget = 20020i32,
    /// Identifies the event that is raised when the specified input reached an element other than the element for which the StartListening method was called.
    InputReachedOtherElement = 20021i32,
    /// Identifies the event that is raised when the specified input was discarded or otherwise failed to reach any element.
    InputDiscarded = 20022i32,
    /// Identifies the event that is raised when a provider issues a system alert. 
    /// 
    /// Supported starting with Windows 8.
    SystemAlert = 20023i32,
    /// Identifies the event that is raised when the content of a live region has changed. 
    /// 
    /// Supported starting with Windows 8.
    LiveRegionChanged = 20024i32,
    /// Identifies the event that is raised when a change is made to the root node of a UI Automation fragment that is hosted in another element. 
    /// 
    /// Supported starting with Windows 8.
    HostedFragmentRootsInvalidated = 20025i32,
    /// Identifies the event that is raised when the user starts to drag an element. This event is raised by the element being dragged. 
    /// 
    /// Supported starting with Windows 8.
    Drag_DragStart = 20026i32,
    /// Identifies the event that is raised when the user ends a drag operation before dropping an element on a drop target. This event is raised by the element being dragged. 
    /// 
    /// Supported starting with Windows 8.
    Drag_DragCancel = 20027i32,
    /// Identifies the event that is raised when the user drops an element on a drop target. This event is raised by the element being dragged. 
    /// Supported starting with Windows 8.
    Drag_DragComplete = 20028i32,
    /// Identifies the event that is raised when the user drags an element into a drop target's boundary. This event is raised by the drop target element. 
    /// 
    /// Supported starting with Windows 8.
    DropTarget_DragEnter = 20029i32,
    /// Identifies the event that is raised when the user drags an element out of a drop target's boundary. This event is raised by the drop target element. 
    /// 
    /// Supported starting with Windows 8.
    DropTarget_DragLeave = 20030i32,
    /// Identifies the event that is raised when the user drops an element on a drop target. This event is raised by the drop target element. 
    /// 
    /// Supported starting with Windows 8.
    DropTarget_Dropped = 20031i32,
    /// Identifies the event that is raised whenever text auto-correction is performed by a control. 
    /// 
    /// Supported starting with Windows 8.1.
    TextEdit_TextChanged = 20032i32,
    /// Identifies the event that is raised whenever a composition replacement is performed by a control. 
    /// 
    /// Supported starting with Windows 8.1.
    TextEdit_ConversionTargetChanged = 20033i32,
    /// Identifies the event that is raised when a provider calls the UiaRaiseChangesEvent function.
    Changes = 20034i32,
    /// Identifies the event that is raised when a provider calls the UiaRaiseNotificationEvent method.
    Notification = 20035i32,
    /// Identifies the event that is raised when the active text position changes, indicated by a navigation event within or between read-only text elements 
    /// (such as web browsers, PDF documents, or EPUB documents) using bookmarks (fragment identifiers that refer to a location within a resource).
    ActiveTextPositionChanged = 20036i32,
}

/// `StructureChangeType` is an enum wrapper for `windows::Win32::UI::Accessibility::StructureChangeType`.
/// 
/// Contains values that specify the type of change in the Microsoft UI Automation tree structure.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::StructureChangeType)]
pub enum StructureChangeType {
    /// A child element was added to the UI Automation element tree.
    ChildAdded = 0i32,
    /// A child element was removed from the UI Automation element tree.
    ChildRemoved = 1i32,
    /// Child elements were invalidated in the UI Automation element tree. 
    /// This might mean that one or more child elements were added or removed, or a combination of both. 
    /// This value can also indicate that one subtree in the UI was substituted for another. 
    /// For example, the entire contents of a dialog box changed at once, or the view of a list changed because an Explorer-type application navigated to another location. 
    /// The exact meaning depends on the UI Automation provider implementation.
    ChildrenInvalidated = 2i32,
    /// Child elements were added in bulk to the UI Automation element tree.
    ChildrenBulkAdded = 3i32,
    /// Child elements were removed in bulk from the UI Automation element tree.
    ChildrenBulkRemoved = 4i32,
    /// The order of child elements has changed in the UI Automation element tree. Child elements may or may not have been added or removed.
    ChildrenReordered = 5i32
}

/// A wrapper for windows `IUIAutomationEventHandler` interface. 
/// 
/// Exposes a method to handle Microsoft UI Automation events.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UIEventHandler {
    handler: IUIAutomationEventHandler
}

impl UIEventHandler {
    /// Handles a Microsoft UI Automation event.
    pub fn handle_automation_event(&self, sender: &UIElement, event: UIEventType) -> Result<()> {
        unsafe {
            self.handler.HandleAutomationEvent(sender, event.into())?;
        }
        Ok(())
    }
}

impl From<IUIAutomationEventHandler> for UIEventHandler {
    fn from(handler: IUIAutomationEventHandler) -> Self {
        Self{
            handler
        }
    }
}

impl From<&IUIAutomationEventHandler> for UIEventHandler {
    fn from(value: &IUIAutomationEventHandler) -> Self {
        value.clone().into()
    }
}

impl Into<IUIAutomationEventHandler> for UIEventHandler {
    fn into(self) -> IUIAutomationEventHandler {
        self.handler
    }
}

impl AsRef<IUIAutomationEventHandler> for UIEventHandler {
    fn as_ref(&self) -> &IUIAutomationEventHandler {
        &self.handler
    }
}

impl Param<IUIAutomationEventHandler> for UIEventHandler {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationEventHandler> {
        // windows::core::ParamValue::Owned(self.handler)
        self.handler.param()
    }
}

impl Param<IUIAutomationEventHandler> for &UIEventHandler {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationEventHandler> {
        // windows::core::ParamValue::Borrowed(self.handler.as_raw())
        (&self.handler).param()
    }
}

/// A wrapper for windows `IUIAutomationPropertyChangedEventHandler` interface. 
/// 
/// Exposes a method to handle Microsoft UI Automation events that occur when a property is changed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UIPropertyChangedEventHandler {
    handler: IUIAutomationPropertyChangedEventHandler
}

impl UIPropertyChangedEventHandler {
    /// Handles a Microsoft UI Automation property-changed event.
    pub fn handle_property_changed_event(&self, sender: &UIElement, property_id: UIProperty, new_value: Variant) -> Result<()> {
        unsafe {
            self.handler.HandlePropertyChangedEvent(sender, property_id.into(), new_value)?
        };
        Ok(())
    }
}

impl From<IUIAutomationPropertyChangedEventHandler> for UIPropertyChangedEventHandler {
    fn from(handler: IUIAutomationPropertyChangedEventHandler) -> Self {
        Self { 
            handler
        }
    }
}

impl From<&IUIAutomationPropertyChangedEventHandler> for UIPropertyChangedEventHandler {
    fn from(value: &IUIAutomationPropertyChangedEventHandler) -> Self {
        value.clone().into()
    }
}

impl Into<IUIAutomationPropertyChangedEventHandler> for UIPropertyChangedEventHandler {
    fn into(self) -> IUIAutomationPropertyChangedEventHandler {
        self.handler
    }
}

impl AsRef<IUIAutomationPropertyChangedEventHandler> for UIPropertyChangedEventHandler {
    fn as_ref(&self) -> &IUIAutomationPropertyChangedEventHandler {
        &self.handler
    }
}

impl Param<IUIAutomationPropertyChangedEventHandler> for UIPropertyChangedEventHandler {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationPropertyChangedEventHandler> {
        self.handler.param()
    }
}

impl Param<IUIAutomationPropertyChangedEventHandler> for &UIPropertyChangedEventHandler {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationPropertyChangedEventHandler> {
        (&self.handler).param()
    }
}

/// A wrapper for windows `IUIAutomationStructureChangedEventHandler` interface. 
/// 
/// Handles an event that is raised when the Microsoft UI Automation tree structure has changed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UIStructureChangeEventHandler {
    handler: IUIAutomationStructureChangedEventHandler
}

impl UIStructureChangeEventHandler {
    /// Handles an event that is raised when the Microsoft UI Automation tree structure has changed.
    pub fn handle_structure_changed_event(&self, sender: &UIElement, change_type: StructureChangeType, runtime_id: Option<&[i32]>) -> Result<()> {
        let runtime_id = if let Some(arr) = runtime_id {
            arr.try_into()?
        } else {
            SafeArray::default()
        };

        unsafe {
            self.handler.HandleStructureChangedEvent(sender, change_type.into(), runtime_id.get_array())?
        }

        Ok(())
    }
}

impl From<IUIAutomationStructureChangedEventHandler> for UIStructureChangeEventHandler {
    fn from(handler: IUIAutomationStructureChangedEventHandler) -> Self {
        Self { 
            handler 
        }
    }
}

impl From<&IUIAutomationStructureChangedEventHandler> for UIStructureChangeEventHandler {
    fn from(value: &IUIAutomationStructureChangedEventHandler) -> Self {
        value.clone().into()
    }
}

impl Into<IUIAutomationStructureChangedEventHandler> for UIStructureChangeEventHandler {
    fn into(self) -> IUIAutomationStructureChangedEventHandler {
        self.handler
    }
}

impl AsRef<IUIAutomationStructureChangedEventHandler> for UIStructureChangeEventHandler {
    fn as_ref(&self) -> &IUIAutomationStructureChangedEventHandler {
        &self.handler
    }
}

impl Param<IUIAutomationStructureChangedEventHandler> for UIStructureChangeEventHandler {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationStructureChangedEventHandler> {
        self.handler.param()
    }
}

impl Param<IUIAutomationStructureChangedEventHandler> for &UIStructureChangeEventHandler {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationStructureChangedEventHandler> {
        (&self.handler).param()
    }
}

/// A wrapper for windows `IUIAutomationFocusChangedEventHandler` interface. 
/// 
/// Exposes a method to handle events that are raised when the keyboard focus moves to another UI Automation element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UIFocusChangedEventHandler {
    handler: IUIAutomationFocusChangedEventHandler
}

impl UIFocusChangedEventHandler {
    /// Handles the event raised when the keyboard focus moves to a different UI Automation element.
    pub fn handle_focus_changed_event(&self, sender: &UIElement) -> Result<()> {
        unsafe {
            self.handler.HandleFocusChangedEvent(sender)?
        };
        Ok(())
    }
}

impl From<IUIAutomationFocusChangedEventHandler> for UIFocusChangedEventHandler {
    fn from(handler: IUIAutomationFocusChangedEventHandler) -> Self {
        Self {
            handler
        }
    }
}

impl From<&IUIAutomationFocusChangedEventHandler> for UIFocusChangedEventHandler {
    fn from(value: &IUIAutomationFocusChangedEventHandler) -> Self {
        value.clone().into()
    }
}

impl Into<IUIAutomationFocusChangedEventHandler> for UIFocusChangedEventHandler {
    fn into(self) -> IUIAutomationFocusChangedEventHandler {
        self.handler
    }
}

impl AsRef<IUIAutomationFocusChangedEventHandler> for UIFocusChangedEventHandler {
    fn as_ref(&self) -> &IUIAutomationFocusChangedEventHandler {
        &self.handler
    }
}

impl Param<IUIAutomationFocusChangedEventHandler> for UIFocusChangedEventHandler {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationFocusChangedEventHandler> {
        self.handler.param()
    }
}

impl Param<IUIAutomationFocusChangedEventHandler> for &UIFocusChangedEventHandler {
    unsafe fn param(self) -> windows::core::ParamValue<IUIAutomationFocusChangedEventHandler> {
        (&self.handler).param()
    }
}

/// Defines a custom handler for `IUIAutomationEventHandler`.
pub trait CustomEventHandler {
    fn handle(&self, sender: &UIElement, event_type: UIEventType) -> Result<()>;
}

#[implement(IUIAutomationEventHandler)]
pub struct AutomationEventHandler {
    handler: Box<dyn CustomEventHandler>
}

impl IUIAutomationEventHandler_Impl for AutomationEventHandler {
    fn HandleAutomationEvent(&self, sender: Option<&IUIAutomationElement>, eventid: UIA_EVENT_ID) -> windows::core::Result<()> {
        if let Some(e) = sender {
            let element = UIElement::from(e);
            self.handler.handle(&element, eventid.into()).map_err(|e| e.into())
        } else {
            Ok(())
        }
    }
}

impl <T> From<T> for AutomationEventHandler where T: CustomEventHandler + 'static {
    fn from(value: T) -> Self {
        Self {
            handler: Box::new(value)
        }
    }
}

impl <T> From<T> for UIEventHandler where T: CustomEventHandler + 'static {
    fn from(value: T) -> Self {
        let handler = AutomationEventHandler::from(value);
        let handler: IUIAutomationEventHandler = handler.into();
        handler.into()
    }
}

#[cfg(test)]
mod tests {
    use windows::Win32::UI::Accessibility::UIA_DropTarget_DroppedEventId;

    use crate::UIAutomation;

    use super::CustomEventHandler;
    use super::UIEventHandler;
    use super::UIEventType;

    #[test]
    fn test_uievent_types() {
        let t = UIEventType::try_from(UIA_DropTarget_DroppedEventId.0).unwrap();
        assert_eq!(t, UIEventType::DropTarget_Dropped);
    }

    struct MyEventHandler {
    }

    impl CustomEventHandler for MyEventHandler {
        fn handle(&self, sender: &crate::UIElement, event_type: UIEventType) -> crate::Result<()> {
            println!("event: {:?}, element: {}", event_type, sender);
            Ok(())
        }
    }

    #[test]
    fn test_event_handler_trait() {
        let automation = UIAutomation::new().unwrap();

        let root = automation.get_root_element().unwrap();
        let handler = UIEventHandler::from(MyEventHandler {});
        automation.add_automation_event_handler(UIEventType::TextEdit_TextChanged, &root, crate::types::TreeScope::Subtree, None, &handler).unwrap();
        automation.remove_automation_event_handler(UIEventType::TextEdit_TextChanged, &root, &handler).unwrap();
    }
}