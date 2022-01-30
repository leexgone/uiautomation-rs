use windows::Win32::UI::Accessibility::IUIAutomationElement;

pub trait Condition {
    fn judge(&self, element: IUIAutomationElement) -> bool;
}