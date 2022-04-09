use windows::Win32::UI::Accessibility::IUIAutomationElement;
use windows::Win32::UI::Accessibility::UIA_ButtonControlTypeId;
use windows::Win32::UI::Accessibility::UIA_InvokePatternId;
use windows::Win32::UI::Accessibility::UIA_ListItemControlTypeId;
use windows::Win32::UI::Accessibility::UIA_SelectionItemPatternId;

use crate::Error;
use crate::Result;
use crate::UIElement;
use crate::errors::ERR_TYPE;
use crate::patterns::UIInvokePattern;
use crate::patterns::UIScrollItemPattern;
use crate::patterns::UISelectionItemPattern;

/// Wrapper an button element as a control.
pub struct ButtonControl {
    control: UIElement
}

impl ButtonControl {
    /// Perform a click event on this control.
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
            Err(Error::new(ERR_TYPE, "Error Control Type"))
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

/// Wrapper a list element as a control. The control type of the element must be `UIA_ListControlTypeId`.
pub struct ListControl {
    control: UIElement
}

/// Wrapper a listitem element as a control. The control type of the element must be `UIA_ListItemControlTypeId`.
pub struct ListItemControl {
    control: UIElement
}

impl ListItemControl {
    /// Determines whether this control is clickable.
    pub fn is_clickable(&self) -> bool {
        let element: &IUIAutomationElement = self.control.as_ref();
        let pattern = unsafe { 
            element.GetCurrentPattern(UIA_InvokePatternId)
        };
        pattern.is_ok()
    }

    /// Perform a click event on this control.
    pub fn click(&self) -> Result<()> {
        let pattern: UIInvokePattern = self.control.get_pattern()?;
        pattern.invoke()
    }

    /// Scroll this item into view.
    pub fn scroll_into_view(&self) -> Result<()> {
        let pattern: UIScrollItemPattern = self.control.get_pattern()?;
        pattern.scroll_into_view()
    }

    /// Determines whether this control is slectable.
    pub fn is_selectable(&self) -> bool {
        let element: &IUIAutomationElement = self.control.as_ref();
        let pattern = unsafe {
            element.GetCurrentPattern(UIA_SelectionItemPatternId)
        };
        pattern.is_ok()
    }

    /// Select current item.
    pub fn select(&self) -> Result<()> {
        let pattern: UISelectionItemPattern = self.control.get_pattern()?;
        pattern.select()
    }

    /// Add current item to selection.
    pub fn add_to_selection(&self) -> Result<()> {
        let pattern: UISelectionItemPattern = self.control.get_pattern()?;
        pattern.add_to_selection()
    }

    /// Remove current item from selection.
    pub fn remove_from_selection(&self) -> Result<()> {
        let pattern: UISelectionItemPattern = self.control.get_pattern()?;
        pattern.remove_from_selection()
    }

    /// Determines whether this item is selected.
    pub fn is_selected(&self) -> Result<bool> {
        let pattern: UISelectionItemPattern = self.control.get_pattern()?;
        pattern.is_selected()
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
            Err(Error::new(ERR_TYPE, "Error Control Type"))
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