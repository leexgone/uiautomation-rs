use windows::Win32::UI::Accessibility::IUIAutomationElement;
use windows::Win32::UI::Accessibility::IUIAutomationEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationEventHandler_Impl;
use windows::Win32::UI::Accessibility::IUIAutomationFocusChangedEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationFocusChangedEventHandler_Impl;
use windows::Win32::UI::Accessibility::IUIAutomationPropertyChangedEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationPropertyChangedEventHandler_Impl;
use windows::Win32::UI::Accessibility::IUIAutomationStructureChangedEventHandler;
use windows::Win32::UI::Accessibility::IUIAutomationStructureChangedEventHandler_Impl;
use windows::Win32::UI::Accessibility::UIA_EVENT_ID;
use windows::Win32::UI::Accessibility::UIA_PROPERTY_ID;
use windows_core::implement;

use crate::types::StructureChangeType;
use crate::variants::SafeArray;
use crate::variants::Variant;
use crate::UIElement;

use super::CustomEventHandler;
use super::CustomFocusChangedEventHandler;
use super::CustomPropertyChangedEventHandler;
use super::CustomStructureChangedEventHandler;

#[implement(IUIAutomationEventHandler)]
pub struct AutomationEventHandler {
    handler: Box<dyn CustomEventHandler>
}

impl IUIAutomationEventHandler_Impl for AutomationEventHandler_Impl {
    fn HandleAutomationEvent(&self, sender: windows_core::Ref<'_, IUIAutomationElement>, eventid: UIA_EVENT_ID) -> windows::core::Result<()> {
        if let Some(e) = sender.as_ref() { 
            let element = UIElement::from(e);
            match eventid.try_into() {
                Ok(event_id) => self.handler.handle(&element, event_id).map_err(|e| e.into()),
                Err(e) => Err(e.into())
            }
            // self.handler.handle(&element, eventid.into()).map_err(|e| e.into())
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

impl IUIAutomationPropertyChangedEventHandler_Impl for AutomationPropertyChangedHandler_Impl {
    fn HandlePropertyChangedEvent(&self, sender: windows_core::Ref<'_, IUIAutomationElement>, propertyid: UIA_PROPERTY_ID, newvalue: &windows::Win32::System::Variant::VARIANT) -> windows::core::Result<()> {
        if let Some(e) = sender.as_ref() {
            let element = UIElement::from(e);
            let value = Variant::from(newvalue);
            match propertyid.try_into() {
                Ok(property_id) => self.handler.handle(&element, property_id, value).map_err(|e| e.into()),
                Err(e) => Err(e.into()),
            }
            // self.handler.handle(&element, propertyid.into(), value).map_err(|e| e.into())
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
pub struct AutomationStructureChangedEventHandler {
    handler: Box<dyn CustomStructureChangedEventHandler>
}

impl IUIAutomationStructureChangedEventHandler_Impl for AutomationStructureChangedEventHandler_Impl {
    fn HandleStructureChangedEvent(&self, sender: windows_core::Ref<'_, IUIAutomationElement>, changetype: windows::Win32::UI::Accessibility::StructureChangeType, runtimeid: *const windows::Win32::System::Com::SAFEARRAY) -> windows_core::Result<()> {
        if let Some(e) = sender.as_ref() {
            let element = UIElement::from(e);
            let arr = SafeArray::from(runtimeid);
            let change_type: StructureChangeType = match changetype.try_into() {
                Ok(change_type) => change_type,
                Err(e) => return Err(e.into())
            };
            let ret = if arr.is_null() {
                self.handler.handle(&element, change_type, None)
            } else {
                let runtime_id: Vec<i32> = match arr.try_into() {
                    Ok(arr) => arr,
                    Err(e) => return  Err(e.into())
                };
                self.handler.handle(&element, change_type, Some(&runtime_id))
            };
            ret.map_err(|e| e.into())
        } else {
            Ok(())
        }
    }
}

impl <T> From<T> for AutomationStructureChangedEventHandler where T: CustomStructureChangedEventHandler + 'static {
    fn from(value: T) -> Self {
        Self {
            handler: Box::new(value)
        }
    }
}

#[implement(IUIAutomationFocusChangedEventHandler)]
pub struct AutomationFocusChangedEventHandler { 
    handler: Box<dyn CustomFocusChangedEventHandler>
}

impl IUIAutomationFocusChangedEventHandler_Impl for AutomationFocusChangedEventHandler_Impl {
    fn HandleFocusChangedEvent(&self, sender: windows_core::Ref<'_, IUIAutomationElement>) -> windows_core::Result<()> {
        if let Some(e) = sender.as_ref() {
            let element = UIElement::from(e);
            self.handler.handle(&element).map_err(|e| e.into())
        } else {
            Ok(())
        }
    }
}

impl <T> From<T> for AutomationFocusChangedEventHandler where T: CustomFocusChangedEventHandler + 'static {
    fn from(value: T) -> Self {
        Self {
            handler: Box::new(value)
        }
    }
}
