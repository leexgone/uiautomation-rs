use windows::Win32::UI::Accessibility::IUIAutomationElement;
use windows::Win32::UI::Accessibility::IUIAutomationEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationEventHandler_Impl;
use windows::Win32::UI::Accessibility::IUIAutomationPropertyChangedEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationPropertyChangedEventHandler_Impl;
use windows::Win32::UI::Accessibility::IUIAutomationStructureChangedEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationStructureChangedEventHandler_Impl;
use windows::Win32::UI::Accessibility::UIA_EVENT_ID;
use windows::Win32::UI::Accessibility::UIA_PROPERTY_ID;
use windows_core::implement;

use crate::variants::SafeArray;
use crate::variants::Variant;
use crate::UIElement;

use super::CustomEventHandler;
use super::CustomPropertyChangedEventHandler;
use super::CustomStructureChangeEventHandler;

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

#[implement(IUIAutomationStructureChangedEventHandler)]
pub struct AutomationStructureChangeEventHandler {
    handler: Box<dyn CustomStructureChangeEventHandler>
}

impl IUIAutomationStructureChangedEventHandler_Impl for AutomationStructureChangeEventHandler {
    fn HandleStructureChangedEvent(&self, sender: Option<&IUIAutomationElement>, changetype: windows::Win32::UI::Accessibility::StructureChangeType, runtimeid: *const windows::Win32::System::Com::SAFEARRAY) -> windows_core::Result<()> {
        if let Some(e) = sender {
            let element = UIElement::from(e);
            let arr = SafeArray::from(runtimeid);
            let ret = if arr.is_null() {
                self.handler.handle(&element, changetype.into(), None)
            } else {
                let runtime_id: Vec<i32> = match arr.try_into() {
                    Ok(arr) => arr,
                    Err(e) => return  Err(e.into())
                };
                self.handler.handle(&element, changetype.into(), Some(&runtime_id))
            };
            ret.map_err(|e| e.into())
        } else {
            Ok(())
        }
    }
}

impl <T> From<T> for AutomationStructureChangeEventHandler where T: CustomStructureChangeEventHandler + 'static {
    fn from(value: T) -> Self {
        Self {
            handler: Box::new(value)
        }
    }
}
