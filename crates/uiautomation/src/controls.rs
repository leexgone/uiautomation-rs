use std::fmt::Display;

use uiautomation_derive::*;
use windows::Win32::UI::Accessibility::*;
use windows::Win32::UI::WindowsAndMessaging::SetForegroundWindow;

use super::actions::*;
use super::Error;
use super::Result;
use super::UIElement;
use super::errors::ERR_TYPE;
use super::patterns::*;

macro_rules! as_control {
    ($control: ident) => {
        if $control.get_control_type()? == Self::TYPE_ID { //$type_id {
            Ok(Self {
                control: $control
            })
        } else {
            Err(Error::new(ERR_TYPE, "Error Control Type"))
        }
    };
}

macro_rules! as_control_ref {
    ($control: ident) => {
        if $control.get_control_type()? == Self::TYPE_ID { 
            Ok(Self {
                control: $control.clone()
            })
        } else {
            Err(Error::new(ERR_TYPE, "Error Control Type"))
        }
    };
}

/// `Control` is the trait for ui element control.
pub trait Control {
    /// Defines the control type id.
    const TYPE_ID: UIA_CONTROLTYPE_ID;
}

/// Wrapper a AppBar element as control. The control type of the element must be `UIA_AppBarControlTypeId`.
/// 
/// + Must support: `ExpandCollapse`, `Toggle`
/// + Conditional support: None
#[derive(Debug, ExpandCollapse, Toggle)]
pub struct AppBarControl {
    control: UIElement
}

impl Control for AppBarControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_AppBarControlTypeId;
}

impl TryFrom<UIElement> for AppBarControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for AppBarControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
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

/// Wrapper a Button element as a control. The control type of the element must be `UIA_ButtonControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support： `ExpandCollapse`, `Invoke`, `Toggle`, `Value`
#[derive(Debug, Invoke, Value, ExpandCollapse, Toggle)]
pub struct ButtonControl {
    control: UIElement
}

impl Control for ButtonControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_ButtonControlTypeId;
}

impl TryFrom<UIElement> for ButtonControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for ButtonControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
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

impl Control for CalendarControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_CalendarControlTypeId;
}

impl TryFrom<UIElement> for CalendarControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for CalendarControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
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

/// Wrapper a CheckBox element as control. The control type of the element must be `UIA_CheckBoxControlTypeId`.
/// 
/// + Must support: `Toggle`
/// + Conditional support: None
#[derive(Debug, Toggle)]
pub struct CheckBoxControl {
    control: UIElement
}

impl Control for CheckBoxControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_CheckBoxControlTypeId;
}

impl TryFrom<UIElement> for CheckBoxControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for CheckBoxControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for CheckBoxControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for CheckBoxControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for CheckBoxControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "CheckBox({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a ComboBox element as control. The control type of the element must be `UIA_ComboBoxControlTypeId`.
/// 
/// + Must support: `ExpandCollapse`
/// + Conditional support: `Selection`, `Value`
#[derive(Debug, ExpandCollapse, Selection, Value)]
pub struct ComboBoxControl {
    control: UIElement
}

impl Control for ComboBoxControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_ComboBoxControlTypeId;
}

impl TryFrom<UIElement> for ComboBoxControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for ComboBoxControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for ComboBoxControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for ComboBoxControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for ComboBoxControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "ComboBox({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a DataGrid element as control. The control type of the element must be `UIA_DataGridControlTypeId`.
/// 
/// + Must support: `Grid`
/// + Conditional support: `Scroll`, `Selection`, `Table`
#[derive(Debug, Grid, Scroll, Selection, Table)]
pub struct DataGridControl {
    control: UIElement
}

impl Control for DataGridControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_DataGridControlTypeId;
}

impl TryFrom<UIElement> for DataGridControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for DataGridControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for DataGridControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for DataGridControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for DataGridControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "DataGrid({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a DataItem element as control. The control type of the element must be `UIA_DataItemControlTypeId`.
/// 
/// + Must support: `SelectionItem`
/// + Conditional support: `CustomNavigation`, `ExpandCollapse`, `GridItem`, `ScrollItem`, `TableItem`, `Toggle`, `Value`
#[derive(Debug, SelectionItem, CustomNavigation, ExpandCollapse, GridItem, ScrollItem, TableItem, Toggle, Value)]
pub struct DataItemControl {
    control: UIElement
}

impl Control for DataItemControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_DataItemControlTypeId;
}

impl TryFrom<UIElement> for DataItemControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for DataItemControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for DataItemControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for DataItemControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for DataItemControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "DataItem({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Document element as control. The control type of the element must be `UIA_DocumentControlTypeId`.
/// 
/// + Must support: `Text`
/// + Conditional support: `Scroll`, `Value`
#[derive(Debug, Text, Scroll, Value)]
pub struct DocumentControl {
    control: UIElement
}

impl Control for DocumentControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_DocumentControlTypeId;
}

impl TryFrom<UIElement> for DocumentControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for DocumentControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for DocumentControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for DocumentControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for DocumentControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Document({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Edit element as control. The control type of the element must be `UIA_EditControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `RangeValue`, `Text`, `Value`
#[derive(Debug, RangeValue, Text, Value)]
pub struct EditControl {
    control: UIElement
}

impl Control for EditControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_EditControlTypeId;
}

impl TryFrom<UIElement> for EditControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for EditControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
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

/// Wrapper a Group element as control. The control type of the element must be `UIA_GroupControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `ExpandCollapse`
#[derive(Debug, ExpandCollapse)]
pub struct GroupControl {
    control: UIElement
}

impl Control for GroupControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_GroupControlTypeId;
}

impl TryFrom<UIElement> for GroupControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for GroupControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for GroupControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for GroupControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for GroupControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Group({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Header element as control. The control type of the element must be `UIA_HeaderControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Transform`
#[derive(Debug, Transform)]
pub struct HeaderControl {
    control: UIElement
}

impl Control for HeaderControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_HeaderControlTypeId;
}

impl TryFrom<UIElement> for HeaderControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for HeaderControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for HeaderControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for HeaderControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for HeaderControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Header({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a HeaderItem element as control. The control type of the element must be `UIA_HeaderItemControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `CustomNavigation`, `Invoke`, `Transform`
#[derive(Debug, CustomNavigation, Invoke, Transform)]
pub struct HeaderItemControl {
    control: UIElement
}

impl Control for HeaderItemControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_HeaderItemControlTypeId;
}

impl TryFrom<UIElement> for HeaderItemControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for HeaderItemControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for HeaderItemControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for HeaderItemControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for HeaderItemControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "HeaderItem({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Hyperlink element as control. The control type of the element must be `UIA_HyperlinkControlTypeId`.
/// 
/// + Must support: `Invoke`
/// + Conditional support: `Value`
#[derive(Debug, Invoke, Value)]
pub struct HyperlinkControl {
    control: UIElement
}

impl Control for HyperlinkControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_HyperlinkControlTypeId;
}

impl TryFrom<UIElement> for HyperlinkControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for HyperlinkControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for HyperlinkControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for HyperlinkControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for HyperlinkControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Hyperlink({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Image element as control. The control type of the element must be `UIA_ImageControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `GridItem`, `TableItem`
#[derive(Debug, GridItem, TableItem)]
pub struct ImageControl {
    control: UIElement
}

impl Control for ImageControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_ImageControlTypeId;
}

impl TryFrom<UIElement> for ImageControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for ImageControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for ImageControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for ImageControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for ImageControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Image({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a List element as control. The control type of the element must be `UIA_ListControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Grid`, `MultipleView`, `Scroll`, `Selection`
#[derive(Debug, Grid, MultipleView, Scroll, Selection)]
pub struct ListControl {
    control: UIElement
}

impl Control for ListControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_ListControlTypeId;
}

impl TryFrom<UIElement> for ListControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for ListControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
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

/// Wrapper a ListItem element as control. The control type of the element must be `UIA_ListItemControlTypeId`.
/// 
/// + Must support: `SelectionItem`
/// + Conditional support: 	`CustomNavigation`, `ExpandCollapse`, `GridItem`, `Invoke`, `ScrollItem`, `Toggle`, `Value`
#[derive(Debug, SelectionItem, CustomNavigation, ExpandCollapse, GridItem, Invoke, ScrollItem, Toggle, Value)]
pub struct ListItemControl {
    control: UIElement
}

impl Control for ListItemControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_ListItemControlTypeId;
}

impl TryFrom<UIElement> for ListItemControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for ListItemControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
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

/// Wrapper a Menu element as control. The control type of the element must be `UIA_MenuControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: None
#[derive(Debug)]
pub struct MenuControl {
    control: UIElement
}

impl Control for MenuControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_MenuControlTypeId;
}

impl TryFrom<UIElement> for MenuControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for MenuControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for MenuControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for MenuControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for MenuControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Menu({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a MenuBar element as control. The control type of the element must be `UIA_MenuBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Dock`, `ExpandCollapse`, `Transform`
#[derive(Debug, Dock, ExpandCollapse, Transform)]
pub struct MenuBarControl {
    control: UIElement
}

impl Control for MenuBarControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_MenuBarControlTypeId;
}

impl TryFrom<UIElement> for MenuBarControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for MenuBarControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for MenuBarControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for MenuBarControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for MenuBarControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "MenuBar({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a MenuItem element as control. The control type of the element must be `UIA_MenuItemControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `ExpandCollapse`, `Invoke`, `SelectionItem`, `Toggle`
#[derive(Debug, ExpandCollapse, Invoke, SelectionItem, Toggle)]
pub struct MenuItemControl {
    control: UIElement
}

impl Control for MenuItemControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_MenuItemControlTypeId;
}

impl TryFrom<UIElement> for MenuItemControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for MenuItemControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for MenuItemControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for MenuItemControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for MenuItemControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "MenuItem({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Pane element as control. The control type of the element must be `UIA_PaneControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Dock`, `Scroll`, `Transform`
#[derive(Debug,	Dock, Scroll, Transform)]
pub struct PaneControl {
    control: UIElement
}

impl Control for PaneControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_PaneControlTypeId;
}

impl TryFrom<UIElement> for PaneControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for PaneControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for PaneControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for PaneControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for PaneControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Pane({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a ProgressBar element as control. The control type of the element must be `UIA_ProgressBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: RangeValue, Value
#[derive(Debug, RangeValue, Value)]
pub struct ProgressBarControl {
    control: UIElement
}

impl Control for ProgressBarControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_ProgressBarControlTypeId;
}

impl TryFrom<UIElement> for ProgressBarControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for ProgressBarControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for ProgressBarControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for ProgressBarControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for ProgressBarControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "ProgressBar({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a RadioButton element as control. The control type of the element must be `UIA_RadioButtonControlTypeId`.
/// 
/// + Must support: `SelectionItem`
/// + Conditional support: None
#[derive(Debug, SelectionItem)]
pub struct RadioButtonControl {
    control: UIElement
}

impl Control for RadioButtonControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_RadioButtonControlTypeId;
}

impl TryFrom<UIElement> for RadioButtonControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for RadioButtonControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for RadioButtonControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for RadioButtonControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for RadioButtonControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "RadioButton({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a ScrollBar element as control. The control type of the element must be `UIA_ScrollBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `RangeValue`
#[derive(Debug, RangeValue)]
pub struct ScrollBarControl {
    control: UIElement
}

impl Control for ScrollBarControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_ScrollBarControlTypeId;
}

impl TryFrom<UIElement> for ScrollBarControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for ScrollBarControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for ScrollBarControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for ScrollBarControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for ScrollBarControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "ScrollBar({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a SemanticZoom element as control. The control type of the element must be `UIA_SemanticZoomControlTypeId`.
/// 
/// + Must support: `Toggle`
/// + Conditional support: None
#[derive(Debug, Toggle)]
pub struct SemanticZoomControl {
    control: UIElement
}

impl Control for SemanticZoomControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_SemanticZoomControlTypeId;
}

impl TryFrom<UIElement> for SemanticZoomControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for SemanticZoomControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for SemanticZoomControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for SemanticZoomControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for SemanticZoomControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "SemanticZoom({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Separator element as control. The control type of the element must be `UIA_SeparatorControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: None
#[derive(Debug)]
pub struct SeparatorControl {
    control: UIElement
}

impl Control for SeparatorControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_SeparatorControlTypeId;
}

impl TryFrom<UIElement> for SeparatorControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for SeparatorControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for SeparatorControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for SeparatorControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for SeparatorControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Separator({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Slider element as control. The control type of the element must be `UIA_SliderControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `RangeValue`, `Selection`, `Value`
#[derive(Debug, RangeValue, Selection, Value)]
pub struct SliderControl {
    control: UIElement
}

impl Control for SliderControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_SliderControlTypeId;
}

impl TryFrom<UIElement> for SliderControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for SliderControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for SliderControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for SliderControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for SliderControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Slider({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Spinner element as control. The control type of the element must be `UIA_SpinnerControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `RangeValue`, `Selection`, `Value`
#[derive(Debug, RangeValue, Selection, Value)]
pub struct SpinnerControl {
    control: UIElement
}

impl Control for SpinnerControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_SpinnerControlTypeId;
}

impl TryFrom<UIElement> for SpinnerControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for SpinnerControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for SpinnerControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for SpinnerControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for SpinnerControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Spinner({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a SplitButton element as control. The control type of the element must be `UIA_SplitButtonControlTypeId`.
/// 
/// + Must support: `ExpandCollapse`, `Invoke`
/// + Conditional support: None
#[derive(Debug, ExpandCollapse, Invoke)]
pub struct SplitButtonControl {
    control: UIElement
}

impl Control for SplitButtonControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_SplitButtonControlTypeId;
}

impl TryFrom<UIElement> for SplitButtonControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for SplitButtonControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for SplitButtonControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for SplitButtonControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for SplitButtonControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "SplitButton({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a StatusBar element as control. The control type of the element must be `UIA_StatusBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Grid`
#[derive(Debug, Grid)]
pub struct StatusBarControl {
    control: UIElement
}

impl Control for StatusBarControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_StatusBarControlTypeId;
}

impl TryFrom<UIElement> for StatusBarControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for StatusBarControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for StatusBarControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for StatusBarControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for StatusBarControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "StatusBar({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Tab element as control. The control type of the element must be `UIA_TabControlTypeId`.
/// 
/// + Must support: `Selection`
/// + Conditional support: `Scroll`
#[derive(Debug, Selection, Scroll)]
pub struct TabControl {
    control: UIElement
}

impl Control for TabControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_TabControlTypeId;
}

impl TryFrom<UIElement> for TabControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for TabControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for TabControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for TabControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for TabControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Tab({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a TabItem element as control. The control type of the element must be `UIA_TabItemControlTypeId`.
/// 
/// + Must support: `SelectionItem`
/// + Conditional support: None
#[derive(Debug, SelectionItem)]
pub struct TabItemControl {
    control: UIElement
}

impl Control for TabItemControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_TabItemControlTypeId;
}

impl TryFrom<UIElement> for TabItemControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for TabItemControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for TabItemControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for TabItemControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for TabItemControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "TabItem({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Table element as control. The control type of the element must be `UIA_TableControlTypeId`.
/// 
/// + Must support: `Grid`, `GridItem`, `Table`, `TableItem`
/// + Conditional support: None
#[derive(Debug, Grid, GridItem, Table, TableItem)]
pub struct TableControl {
    control: UIElement
}

impl Control for TableControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_TableControlTypeId;
}

impl TryFrom<UIElement> for TableControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for TableControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for TableControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for TableControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for TableControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Table({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Text element as control. The control type of the element must be `UIA_TextControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `GridItem`, `TableItem`, `Text`
#[derive(Debug, GridItem, TableItem, Text)]
pub struct TextControl {
    control: UIElement
}

impl Control for TextControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_TextControlTypeId;
}

impl TryFrom<UIElement> for TextControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for TextControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for TextControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for TextControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for TextControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Text({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Thumb element as control. The control type of the element must be `UIA_ThumbControlTypeId`.
/// 
/// + Must support: `Transform`
/// + Conditional support: None
#[derive(Debug, Transform)]
pub struct ThumbControl {
    control: UIElement
}

impl Control for ThumbControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_ThumbControlTypeId;
}

impl TryFrom<UIElement> for ThumbControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for ThumbControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for ThumbControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for ThumbControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for ThumbControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Thumb({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a TitleBar element as control. The control type of the element must be `UIA_TitleBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: None
#[derive(Debug)]
pub struct TitleBarControl {
    control: UIElement
}

impl Control for TitleBarControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_TitleBarControlTypeId;
}

impl TryFrom<UIElement> for TitleBarControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for TitleBarControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for TitleBarControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for TitleBarControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for TitleBarControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "TitleBar({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a ToolBar element as control. The control type of the element must be `UIA_ToolBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Dock`, `ExpandCollapse`, `Transform`
#[derive(Debug, Dock, ExpandCollapse, Transform)]
pub struct ToolBarControl {
    control: UIElement
}

impl Control for ToolBarControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_ToolBarControlTypeId;
}

impl TryFrom<UIElement> for ToolBarControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for ToolBarControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for ToolBarControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for ToolBarControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for ToolBarControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "ToolBar({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a ToolTip element as control. The control type of the element must be `UIA_ToolTipControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Text`, `Window`
#[derive(Debug, Text, Window)]
pub struct ToolTipControl {
    control: UIElement
}

impl Control for ToolTipControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_ToolTipControlTypeId;
}

impl TryFrom<UIElement> for ToolTipControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for ToolTipControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for ToolTipControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for ToolTipControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for ToolTipControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "ToolTip({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Tree element as control. The control type of the element must be `UIA_TreeControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Scroll`, `Selection`
#[derive(Debug,	Scroll, Selection)]
pub struct TreeControl {
    control: UIElement
}

impl Control for TreeControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_TreeControlTypeId;
}

impl TryFrom<UIElement> for TreeControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for TreeControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for TreeControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for TreeControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for TreeControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Tree({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a TreeItem element as control. The control type of the element must be `UIA_TreeItemControlTypeId`.
/// 
/// + Must support: `ExpandCollapse`
/// + Conditional support: `Invoke`, `ScrollItem`, `SelectionItem`, `Toggle`
#[derive(Debug,	ExpandCollapse, Invoke, ScrollItem, SelectionItem, Toggle)]
pub struct TreeItemControl {
    control: UIElement
}

impl Control for TreeItemControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_TreeItemControlTypeId;
}

impl TryFrom<UIElement> for TreeItemControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for TreeItemControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
    }
}

impl Into<UIElement> for TreeItemControl {
    fn into(self) -> UIElement {
        self.control
    }
}

impl AsRef<UIElement> for TreeItemControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

impl Display for TreeItemControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "TreeItem({})", self.control.get_name().unwrap_or_default())
    }
}

/// Wrapper a Window element as control. The control type of the element must be `UIA_WindowControlTypeId`.
/// 
/// + Must support: `Transform`, `Window`
/// + Conditional support: `Dock`
#[derive(Debug, Transform, Window, Dock)]
pub struct WindowControl {
    control: UIElement
}

impl WindowControl {
    /// Brings the thread that created the specified window into the foreground and activates the window. 
    pub fn set_foregrand(&self) -> Result<bool> {
        let hwnd = self.control.get_native_window_handle()?;
        let ret = unsafe {
            SetForegroundWindow(hwnd)
        };
        Ok(ret.as_bool())
    }
}

impl Control for WindowControl {
    const TYPE_ID: UIA_CONTROLTYPE_ID = UIA_WindowControlTypeId;
}

impl TryFrom<UIElement> for WindowControl {
    type Error = Error;

    fn try_from(control: UIElement) -> Result<Self> {
        as_control!(control)
    }
}

impl TryFrom<&UIElement> for WindowControl {
    type Error = Error;

    fn try_from(control: &UIElement) -> Result<Self> {
        as_control_ref!(control)
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