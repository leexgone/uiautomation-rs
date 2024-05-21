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

use super::CustomEventHandler;
use super::CustomPropertyChangedEventHandler;

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

#[implement(IUIAutomationPropertyChangedEventHandler)]
pub struct AutomationPropertyChangedHandler {
    handler: Box<dyn CustomPropertyChangedEventHandler>
}

impl IUIAutomationPropertyChangedEventHandler_Impl for AutomationPropertyChangedHandler {
    fn HandlePropertyChangedEvent(&self, sender: Option<&IUIAutomationElement>, propertyid: UIA_PROPERTY_ID, newvalue: &windows::core::VARIANT) -> windows::core::Result<()> {
        if let Some(e) = sender {
            let element = UIElement::from(e);
            let value = Variant::from(newvalue);
            self.handler.handle(&element, propertyid.into(), value).map_err(|e| e.into())
        } else {
            Ok(())
        }
    }
}

impl <T> From<T> for AutomationPropertyChangedHandler where T: CustomPropertyChangedEventHandler + 'static {
    fn from(value: T) -> Self {
        Self {
            handler: Box::new(value)
        }
    }
}
