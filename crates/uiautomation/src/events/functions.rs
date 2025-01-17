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

use super::CustomEventHandlerFn;
use super::CustomFocusChangedEventHandlerFn;
use super::CustomPropertyChangedEventHandlerFn;
use super::CustomStructureChangedEventHandlerFn;

#[implement(IUIAutomationEventHandler)]
pub struct AutomationEventHandler {
    handler: Box<CustomEventHandlerFn>
}

impl IUIAutomationEventHandler_Impl for AutomationEventHandler_Impl {
    fn HandleAutomationEvent(&self, sender: windows_core::Ref<'_, IUIAutomationElement>, eventid: UIA_EVENT_ID) -> windows_core::Result<()> {
        if let Some(e) = sender.as_ref() {
            let element = UIElement::from(e);
            let handler = &self.handler;
            match eventid.try_into() {
                Ok(event_id) => handler(&element, event_id).map_err(|e| e.into()),
                Err(e) => Err(e.into())
            }
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

impl IUIAutomationPropertyChangedEventHandler_Impl for AutomationPropertyChangedEventHandler_Impl {
    fn HandlePropertyChangedEvent(&self, sender: windows_core::Ref<'_, IUIAutomationElement>, propertyid: UIA_PROPERTY_ID, newvalue: &windows::Win32::System::Variant::VARIANT) -> windows_core::Result<()> {
        if let Some(e) = sender.as_ref() {
            let element = UIElement::from(e);
            let value = Variant::from(newvalue);
            let handler = &self.handler;
            match propertyid.try_into() {
                Ok(property_id) => handler(&element, property_id, value).map_err(|e| e.into()),
                Err(e) => Err(e.into()),
            }
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

#[implement(IUIAutomationStructureChangedEventHandler)]
pub struct AutomationStructureChangedEventHandler { 
    handler: Box<CustomStructureChangedEventHandlerFn>
}

impl IUIAutomationStructureChangedEventHandler_Impl for AutomationStructureChangedEventHandler_Impl {
    fn HandleStructureChangedEvent(&self, sender: windows_core::Ref<'_, IUIAutomationElement>, changetype: windows::Win32::UI::Accessibility::StructureChangeType, runtimeid: *const windows::Win32::System::Com::SAFEARRAY) -> windows_core::Result<()> {
        if let Some(e) = sender.as_ref() {
            let handler = &self.handler;
            let element = UIElement::from(e);
            let arr = SafeArray::from(runtimeid);
            let change_type: StructureChangeType = match changetype.try_into() {
                Ok(change_type) => change_type,
                Err(e) => return Err(e.into()),
            };
            let ret = if arr.is_null() {
                handler(&element, change_type, None)
            } else {
                let runtime_id: Vec<i32> = match arr.try_into() {
                    Ok(arr) => arr,
                    Err(e) => return  Err(e.into())
                };
                handler(&element, change_type, Some(&runtime_id))
            };
            ret.map_err(|e| e.into())
        } else {
            Ok(())
        }
    }
}

impl From<Box<CustomStructureChangedEventHandlerFn>> for AutomationStructureChangedEventHandler {
    fn from(handler: Box<CustomStructureChangedEventHandlerFn>) -> Self {
        Self {
            handler
        }
    }
}

#[implement(IUIAutomationFocusChangedEventHandler)]
pub struct AutomationFocusChangedEventHandler {
    handler: Box<CustomFocusChangedEventHandlerFn>
}

impl IUIAutomationFocusChangedEventHandler_Impl for AutomationFocusChangedEventHandler_Impl {
    fn HandleFocusChangedEvent(&self, sender: windows_core::Ref<'_, IUIAutomationElement>) -> windows_core::Result<()> {
        if let Some(e) = sender.as_ref() {
            let element = UIElement::from(e);
            let handler = &self.handler;
            handler(&element).map_err(|e| e.into())
        } else {
            Ok(())
        }
    }
}

impl From<Box<CustomFocusChangedEventHandlerFn>> for AutomationFocusChangedEventHandler {
    fn from(handler: Box<CustomFocusChangedEventHandlerFn>) -> Self {
        Self {
            handler
        }
    }
}
