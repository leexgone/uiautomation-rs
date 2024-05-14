use windows::core::*;
use windows::Win32::UI::Accessibility::*;

#[implement(IUIAutomationEventHandler)]
pub struct MyEventHandler {
}

impl IUIAutomationEventHandler_Impl for MyEventHandler {
    fn HandleAutomationEvent(&self,sender:Option<&IUIAutomationElement>,eventid:UIA_EVENT_ID) -> windows::core::Result<()> {
        todo!()
    }
}