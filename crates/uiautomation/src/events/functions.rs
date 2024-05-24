use windows::Win32::UI::Accessibility::IUIAutomationElement;
use windows::Win32::UI::Accessibility::IUIAutomationEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationEventHandler_Impl;
use windows::Win32::UI::Accessibility::IUIAutomationPropertyChangedEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationPropertyChangedEventHandler_Impl;
use windows::Win32::UI::Accessibility::UIA_EVENT_ID;
use windows::Win32::UI::Accessibility::UIA_PROPERTY_ID;
use windows_core::implement;

use crate::variants::Variant;
use crate::UIElement;

use super::CustomEventHandlerFn;
use super::CustomPropertyChangedEventHandlerFn;

#[implement(IUIAutomationEventHandler)]
pub struct AutomationEventHandler {
    handler: Box<CustomEventHandlerFn>
}

impl IUIAutomationEventHandler_Impl for AutomationEventHandler {
    fn HandleAutomationEvent(&self, sender: Option<&IUIAutomationElement>, eventid: UIA_EVENT_ID) -> windows_core::Result<()> {
        if let Some(e) = sender {
            let element = UIElement::from(e);
            let handler = &self.handler;
            handler(&element, eventid.into()).map_err(|e| e.into())
        } else {
            Ok(())
        }
    }
}

impl From<Box<CustomEventHandlerFn>> for AutomationEventHandler {
    fn from(handler: Box<CustomEventHandlerFn>) -> Self {
        Self {
            handler
        }
    }
}

#[implement(IUIAutomationPropertyChangedEventHandler)]
pub struct AutomationPropertyChangedEventHandler {
    handler: Box<CustomPropertyChangedEventHandlerFn>
}

impl IUIAutomationPropertyChangedEventHandler_Impl for AutomationPropertyChangedEventHandler {
    fn HandlePropertyChangedEvent(&self, sender: Option<&IUIAutomationElement>, propertyid: UIA_PROPERTY_ID, newvalue: &windows_core::VARIANT) -> windows_core::Result<()> {
        if let Some(e) = sender {
            let element = UIElement::from(e);
            let value = Variant::from(newvalue);
            let handler = &self.handler;
            handler(&element, propertyid.into(), value).map_err(|e| e.into())
        } else {
            Ok(())
        }
    }
}

impl From<Box<CustomPropertyChangedEventHandlerFn>> for AutomationPropertyChangedEventHandler {
    fn from(handler: Box<CustomPropertyChangedEventHandlerFn>) -> Self {
        Self {
            handler
        }
    }
}
