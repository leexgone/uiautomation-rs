use std::fmt::Display;

use uiautomation_derive::*;
use windows::Win32::UI::Accessibility::*;

use super::actions::*;
use super::Error;
use super::Result;
use super::UIElement;
use super::errors::ERR_TYPE;
use super::patterns::*;
use super::variants::Variant;

macro_rules! as_control {
    ($control: ident, $type_id: ident) => {
        if $control.get_control_type()? == $type_id {
            Ok(Self {
                $control
            })
        } else {
            Err(Error::new(ERR_TYPE, "Error Control Type"))
        }
    };
}

/// Wrapper a AppBar element as control. The control type of the element must be `UIA_AppBarControlTypeId`.
/// 
/// + Must support: `ExpandCollapse`, `Toggle`
/// + Conditional support: None
#[derive(Debug, ExpandCollapse, Toggle)]
pub struct AppBarControl {
    control: UIElement
}

impl TryFrom<UIElement> for AppBarControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control, UIA_AppBarControlTypeId)
    }
}

impl Into<UIElement> for AppBarControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for AppBarControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for AppBarControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "AppBar({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Button element as a control.
/// 
/// + Must support: None
/// + Conditional support： `ExpandCollapse`, `Invoke`, `Toggle`, `Value`
#[derive(Debug, Invoke, Value, ExpandCollapse, Toggle)]
pub struct ButtonControl {
    control: UIElement
}

impl TryFrom<UIElement> for ButtonControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control, UIA_ButtonControlTypeId)
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

/// Wrapper a Calendar element as control.
/// 
/// + Must support: `Grid`, `Table`
/// + Conditional support： `Scroll`, `Selection`
#[derive(Debug, Grid, Table, Scroll, Selection)]
pub struct CalendarControl {
    control: UIElement
}

impl TryFrom<UIElement> for CalendarControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control, UIA_CalendarControlTypeId)
    }
}

impl Into<UIElement> for CalendarControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for CalendarControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for CalendarControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Calendar({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a window element as control. The control type of the element must be `UIA_WindowControlTypeId`.
#[derive(Window, Transform)]
pub struct WindowControl {
    control: UIElement
}

impl TryFrom<UIElement> for WindowControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control, UIA_WindowControlTypeId)
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

/// Wrapper a list element as a control. The control type of the element must be `UIA_ListControlTypeId`.
#[derive(MultipleView, ItemContainer)]
pub struct ListControl {
    control: UIElement
}

impl TryFrom<UIElement> for ListControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control, UIA_ListControlTypeId)
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
        as_control!(control, UIA_ListItemControlTypeId)
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

/// Wrapper a edit element as a control. The control type of the element must be `UIA_EditControlTypeId`.
#[derive(ScrollItem, Value)]
pub struct EditControl {
    control: UIElement
}

impl TryFrom<UIElement> for EditControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control, UIA_EditControlTypeId)
    }
}

impl Into<UIElement> for EditControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for EditControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for EditControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Edit({})", self.control.get_name().unwrap_or_default())
    }
}
