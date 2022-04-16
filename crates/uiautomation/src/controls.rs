use std::fmt::Display;

use uiautomation_derive::Invoke;
use uiautomation_derive::ItemContainer;
use uiautomation_derive::MultipleView;
use uiautomation_derive::ScrollItem;
use uiautomation_derive::SelectionItem;
use uiautomation_derive::Transform;
use uiautomation_derive::Window;
use windows::Win32::UI::Accessibility::UIA_ButtonControlTypeId;
use windows::Win32::UI::Accessibility::UIA_ListControlTypeId;
use windows::Win32::UI::Accessibility::UIA_ListItemControlTypeId;
use windows::Win32::UI::Accessibility::UIA_WindowControlTypeId;

use crate::actions::*;
use crate::Error;
use crate::Result;
use crate::UIElement;
use crate::errors::ERR_TYPE;
use crate::patterns::UIInvokePattern;
use crate::patterns::UIItemContainerPattern;
use crate::patterns::UIMultipleViewPattern;
use crate::patterns::UIScrollItemPattern;
use crate::patterns::UISelectionItemPattern;
use crate::patterns::UITransformPattern;
use crate::patterns::UIWindowPattern;
use crate::variants::Variant;

/// Wrapper a window element as control. The control type of the element must be `UIA_WindowControlTypeId`.
#[derive(Window, Transform)]
pub struct WindowControl {
    control: UIElement
}

impl TryFrom<UIElement> for WindowControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        if control.get_control_type()? == UIA_WindowControlTypeId {
            Ok(Self {
                control
            })
        } else {
            Err(Error::new(ERR_TYPE, "Error Control Type"))
        }
    }
}

impl Into<UIElement> for WindowControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for WindowControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for WindowControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Window({})", self.control.get_name().unwrap_or_default())
    }
}
/// Wrapper an button element as a control.
#[derive(Invoke)]
pub struct ButtonControl {
    control: UIElement
}

impl TryFrom<UIElement> for ButtonControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        if control.get_control_type()? == UIA_ButtonControlTypeId {
            Ok(Self {
                control
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

impl Display for ButtonControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Button({})", self.control.get_name().unwrap_or_default())
    }
}

impl AsRef<UIElement> for ButtonControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

/// Wrapper a list element as a control. The control type of the element must be `UIA_ListControlTypeId`.
#[derive(MultipleView, ItemContainer)]
pub struct ListControl {
    control: UIElement
}

impl TryFrom<UIElement> for ListControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        if control.get_control_type()? == UIA_ListControlTypeId {
            Ok(Self {
                control
            })
        } else {
            Err(Error::new(ERR_TYPE, "Error Control Type"))
        }
    }
}

impl Into<UIElement> for ListControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for ListControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for ListControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "List({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a listitem element as a control. The control type of the element must be `UIA_ListItemControlTypeId`.
#[derive(Invoke, SelectionItem, ScrollItem)]
pub struct ListItemControl {
    control: UIElement
}

impl TryFrom<UIElement> for ListItemControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        if control.get_control_type()? == UIA_ListItemControlTypeId {
            Ok(Self {
                control
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

impl Display for ListItemControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "ListItem({})", self.control.get_name().unwrap_or_default())
    }
}