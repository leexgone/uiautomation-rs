use windows::Win32::UI::Accessibility::UIA_ButtonControlTypeId;
use windows::Win32::UI::Accessibility::UIA_ListItemControlTypeId;

use crate::Error;
use crate::Result;
use crate::UIElement;
use crate::errors::ERR_TYPE;
use crate::patterns::UIInvokePattern;

/// Wrapper a button element as a control.
pub struct ButtonControl {
    control: UIElement
}

impl ButtonControl {
    pub fn click(&self) -> Result<()> {
        let pattern: UIInvokePattern = self.control.get_pattern()?;
        pattern.invoke()
    }
}

impl TryFrom<UIElement> for ButtonControl {
    type Error = Error;

    fn try_from(value: UIElement) -> Result<Self> {
        if value.get_control_type()? == UIA_ButtonControlTypeId {
            Ok(Self {
                control: value
            })
        } else {
            Err(Error::new(ERR_TYPE, "Unmatched Control Type"))
        }
    }
}

impl Into<UIElement> for ButtonControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for ButtonControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

/// Wrapper a listitem element as control.
pub struct ListItemControl {
    control: UIElement
}

impl ListItemControl {
    pub fn click(&self) -> Result<()> {
        let pattern: UIInvokePattern = self.control.get_pattern()?;
        pattern.invoke()
    }
}

impl TryFrom<UIElement> for ListItemControl {
    type Error = Error;

    fn try_from(value: UIElement) -> Result<Self> {
        if value.get_control_type()? == UIA_ListItemControlTypeId {
            Ok(Self {
                control: value
            })
        } else {
            Err(Error::new(ERR_TYPE, "Unmatched Control Type"))
        }
    }
}

impl Into<UIElement> for ListItemControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for ListItemControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}