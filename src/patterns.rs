use windows::Win32::UI::Accessibility::IUIAutomationInvokePattern;
use windows::Win32::UI::Accessibility::UIA_InvokePatternId;
use windows::core::IUnknown;
use windows::core::Interface;

use crate::errors::Error;
use crate::errors::Result;

pub trait UIPattern : Sized {
    fn pattern_id() -> i32;
    fn new(pattern: IUnknown) -> Result<Self>;
}

pub struct UIInvokePattern {
    pattern: IUIAutomationInvokePattern
}

impl UIInvokePattern {
    pub fn invoke(&self) -> Result<()> {
        unsafe {
            self.pattern.Invoke()?;
        }
        Ok(())
    }
}

impl UIPattern for UIInvokePattern {
    fn pattern_id() -> i32 {
        UIA_InvokePatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIInvokePattern::try_from(pattern)
    }
}

impl From<IUIAutomationInvokePattern> for UIInvokePattern {
    fn from(pattern: IUIAutomationInvokePattern) -> Self {
        Self {
            pattern
        }
    }
}

impl TryFrom<IUnknown> for UIInvokePattern {
    type Error = Error;

    fn try_from(pattern: IUnknown) -> core::result::Result<Self, Self::Error> {
        let pattern: IUIAutomationInvokePattern = pattern.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl Into<IUIAutomationInvokePattern> for UIInvokePattern {
    fn into(self) -> IUIAutomationInvokePattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationInvokePattern> for UIInvokePattern {
    fn as_ref(&self) -> &IUIAutomationInvokePattern {
        &self.pattern
    }
}