use windows::Win32::UI::Accessibility::IUIAutomationElement;
use windows::Win32::UI::Accessibility::IUIAutomationEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationEventHandler_Impl;
use windows::Win32::UI::Accessibility::UIA_EVENT_ID;
use windows_core::implement;

use crate::UIElement;

use super::CustomEventHandlerFn;

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
