use std::fmt::{self , Display, Formatter};

use uiautomation_derive::*;
use windows::Win32::UI::WindowsAndMessaging::SetForegroundWindow;

use super::actions::*;
use super::Result;
use super::UIElement;
use super::errors::ERR_TYPE;
use super::patterns::*;

/// Defines enum for `windows::Win32::UI::Accessibility::UIA_CONTROLTYPE_ID`.
/// 
/// Contains the named constants used to identify Microsoft UI Automation control types.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
pub enum ControlType {
    /// Identifies the Button control type.
    Button = 50000i32,
    /// Identifies the Calendar control type.
    Calendar = 50001i32,
    /// Identifies the CheckBox control type.
    CheckBox = 50002i32,
    /// Identifies the ComboBox control type.
    ComboBox = 50003i32,
    /// Identifies the Edit control type.
    Edit = 50004i32,
    /// Identifies the Hyperlink control type.
    Hyperlink = 50005i32,
    /// Identifies the Image control type.
    Image = 50006i32,
    /// Identifies the ListItem control type.
    ListItem = 50007i32,
    /// Identifies the List control type.
    List = 50008i32,
    /// Identifies the Menu control type.
    Menu = 50009i32,
    /// Identifies the MenuBar control type.
    MenuBar = 50010i32,
    /// Identifies the MenuItem control type.
    MenuItem = 50011i32,
    /// Identifies the ProgressBar control type.
    ProgressBar = 50012i32,
    /// Identifies the RadioButton control type.
    RadioButton = 50013i32,
    /// Identifies the ScrollBar control type.
    ScrollBar = 50014i32,
    /// Identifies the Slider control type.
    Slider = 50015i32,
    /// Identifies the Spinner control type.
    Spinner = 50016i32,
    /// Identifies the StatusBar control type.
    StatusBar = 50017i32,
    /// Identifies the Tab control type.
    Tab = 50018i32,
    /// Identifies the TabItem control type.
    TabItem = 50019i32,
    /// Identifies the Text control type.
    Text = 50020i32,
    /// Identifies the ToolBar control type.
    ToolBar = 50021i32,
    /// Identifies the ToolTip control type.
    ToolTip = 50022i32,
    /// Identifies the Tree control type.
    Tree = 50023i32,
    /// Identifies the TreeItem control type.
    TreeItem = 50024i32,
    /// Identifies the Custom control type. For more information, see Custom Properties, Events, and Control Patterns.
    Custom = 50025i32,
    /// Identifies the Group control type.
    Group = 50026i32,
    /// Identifies the Thumb control type.
    Thumb = 50027i32,
    /// Identifies the DataGrid control type.
    DataGrid = 50028i32,
    /// Identifies the DataItem control type.
    DataItem = 50029i32,
    /// Identifies the Document control type.
    Document = 50030i32,
    /// Identifies the SplitButton control type.
    SplitButton = 50031i32,
    /// Identifies the Window control type.
    Window = 50032i32,
    /// Identifies the Pane control type.
    Pane = 50033i32,
    /// Identifies the Header control type.
    Header = 50034i32,
    /// Identifies the HeaderItem control type.
    HeaderItem = 50035i32,
    /// Identifies the Table control type.
    Table = 50036i32,
    /// Identifies the TitleBar control type.
    TitleBar = 50037i32,
    /// Identifies the Separator control type.
    Separator = 50038i32,
    /// Identifies the SemanticZoom control type. Supported starting with Windows 8.
    SemanticZoom = 50039i32,
    /// Identifies the AppBar control type. Supported starting with Windows 8.1.
    AppBar = 50040i32    
}

impl From<windows::Win32::UI::Accessibility::UIA_CONTROLTYPE_ID> for ControlType {
    fn from(value: windows::Win32::UI::Accessibility::UIA_CONTROLTYPE_ID) -> Self {
        value.0.try_into().unwrap()
    }
}

impl Into<windows::Win32::UI::Accessibility::UIA_CONTROLTYPE_ID> for ControlType {
    fn into(self) -> windows::Win32::UI::Accessibility::UIA_CONTROLTYPE_ID {
        windows::Win32::UI::Accessibility::UIA_CONTROLTYPE_ID(self as _)
    }
}

/// `Control` is the trait for ui element control.
pub trait Control {
    /// Defines the control type id.
    const TYPE: ControlType;
}

/// Wrapper a AppBar element as control. The control type of the element must be `UIA_AppBarControlTypeId`.
/// 
/// + Must support: `ExpandCollapse`, `Toggle`
/// + Conditional support: None
#[derive(Debug, Control, ExpandCollapse, Toggle)]
pub struct AppBarControl {
    control: UIElement
}

impl Control for AppBarControl {
    const TYPE: ControlType = ControlType::AppBar;
}

/// Wrapper a Button element as a control. The control type of the element must be `UIA_ButtonControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support： `ExpandCollapse`, `Invoke`, `Toggle`, `Value`
#[derive(Debug, Control, Invoke, Value, ExpandCollapse, Toggle)]
pub struct ButtonControl {
    control: UIElement
}

impl Control for ButtonControl {
    const TYPE: ControlType = ControlType::Button;
}

/// Wrapper a Calendar element as control.
/// 
/// + Must support: `Grid`, `Table`
/// + Conditional support： `Scroll`, `Selection`
#[derive(Debug, Control, Grid, Table, Scroll, Selection)]
pub struct CalendarControl {
    control: UIElement
}

impl Control for CalendarControl {
    const TYPE: ControlType = ControlType::Calendar;
}

/// Wrapper a CheckBox element as control. The control type of the element must be `UIA_CheckBoxControlTypeId`.
/// 
/// + Must support: `Toggle`
/// + Conditional support: None
#[derive(Debug, Control, Toggle)]
pub struct CheckBoxControl {
    control: UIElement
}

impl Control for CheckBoxControl {
    const TYPE: ControlType = ControlType::CheckBox;
}

/// Wrapper a ComboBox element as control. The control type of the element must be `UIA_ComboBoxControlTypeId`.
/// 
/// + Must support: `ExpandCollapse`
/// + Conditional support: `Selection`, `Value`
#[derive(Debug, Control, ExpandCollapse, Selection, Value)]
pub struct ComboBoxControl {
    control: UIElement
}

impl Control for ComboBoxControl {
    const TYPE: ControlType = ControlType::ComboBox;
}

/// Wrapper a DataGrid element as control. The control type of the element must be `UIA_DataGridControlTypeId`.
/// 
/// + Must support: `Grid`
/// + Conditional support: `Scroll`, `Selection`, `Table`
#[derive(Debug, Control, Grid, Scroll, Selection, Table)]
pub struct DataGridControl {
    control: UIElement
}

impl Control for DataGridControl {
    const TYPE: ControlType = ControlType::DataGrid;
}

/// Wrapper a DataItem element as control. The control type of the element must be `UIA_DataItemControlTypeId`.
/// 
/// + Must support: `SelectionItem`
/// + Conditional support: `CustomNavigation`, `ExpandCollapse`, `GridItem`, `ScrollItem`, `TableItem`, `Toggle`, `Value`
#[derive(Debug, Control, SelectionItem, CustomNavigation, ExpandCollapse, GridItem, ScrollItem, TableItem, Toggle, Value)]
pub struct DataItemControl {
    control: UIElement
}

impl Control for DataItemControl {
    const TYPE: ControlType = ControlType::DataItem;
}

/// Wrapper a Document element as control. The control type of the element must be `UIA_DocumentControlTypeId`.
/// 
/// + Must support: `Text`
/// + Conditional support: `Scroll`, `Value`
#[derive(Debug, Control, Text, Scroll, Value)]
pub struct DocumentControl {
    control: UIElement
}

impl Control for DocumentControl {
    const TYPE: ControlType = ControlType::Document;
}

/// Wrapper a Edit element as control. The control type of the element must be `UIA_EditControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `RangeValue`, `Text`, `Value`
#[derive(Debug, Control, RangeValue, Text, Value)]
pub struct EditControl {
    control: UIElement
}

impl Control for EditControl {
    const TYPE: ControlType = ControlType::Edit;
}

/// Wrapper a Group element as control. The control type of the element must be `UIA_GroupControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `ExpandCollapse`
#[derive(Debug, Control, ExpandCollapse)]
pub struct GroupControl {
    control: UIElement
}

impl Control for GroupControl {
    const TYPE: ControlType = ControlType::Group;
}

/// Wrapper a Header element as control. The control type of the element must be `UIA_HeaderControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Transform`
#[derive(Debug, Control, Transform)]
pub struct HeaderControl {
    control: UIElement
}

impl Control for HeaderControl {
    const TYPE: ControlType = ControlType::Header;
}

/// Wrapper a HeaderItem element as control. The control type of the element must be `UIA_HeaderItemControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `CustomNavigation`, `Invoke`, `Transform`
#[derive(Debug, Control, CustomNavigation, Invoke, Transform)]
pub struct HeaderItemControl {
    control: UIElement
}

impl Control for HeaderItemControl {
    const TYPE: ControlType = ControlType::HeaderItem;
}

/// Wrapper a Hyperlink element as control. The control type of the element must be `UIA_HyperlinkControlTypeId`.
/// 
/// + Must support: `Invoke`
/// + Conditional support: `Value`
#[derive(Debug, Control, Invoke, Value)]
pub struct HyperlinkControl {
    control: UIElement
}

impl Control for HyperlinkControl {
    const TYPE: ControlType = ControlType::Hyperlink;
}

/// Wrapper a Image element as control. The control type of the element must be `UIA_ImageControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `GridItem`, `TableItem`
#[derive(Debug, Control, GridItem, TableItem)]
pub struct ImageControl {
    control: UIElement
}

impl Control for ImageControl {
    const TYPE: ControlType = ControlType::Image;
}

/// Wrapper a List element as control. The control type of the element must be `UIA_ListControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Grid`, `MultipleView`, `Scroll`, `Selection`
#[derive(Debug, Control, Grid, MultipleView, Scroll, Selection)]
pub struct ListControl {
    control: UIElement
}

impl Control for ListControl {
    const TYPE: ControlType = ControlType::List;
}

/// Wrapper a ListItem element as control. The control type of the element must be `UIA_ListItemControlTypeId`.
/// 
/// + Must support: `SelectionItem`
/// + Conditional support: 	`CustomNavigation`, `ExpandCollapse`, `GridItem`, `Invoke`, `ScrollItem`, `Toggle`, `Value`
#[derive(Debug, Control, SelectionItem, CustomNavigation, ExpandCollapse, GridItem, Invoke, ScrollItem, Toggle, Value)]
pub struct ListItemControl {
    control: UIElement
}

impl Control for ListItemControl {
    const TYPE: ControlType = ControlType::ListItem;
}

/// Wrapper a Menu element as control. The control type of the element must be `UIA_MenuControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: None
#[derive(Debug, Control)]
pub struct MenuControl {
    control: UIElement
}

impl Control for MenuControl {
    const TYPE: ControlType = ControlType::Menu;
}

/// Wrapper a MenuBar element as control. The control type of the element must be `UIA_MenuBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Dock`, `ExpandCollapse`, `Transform`
#[derive(Debug, Control, Dock, ExpandCollapse, Transform)]
pub struct MenuBarControl {
    control: UIElement
}

impl Control for MenuBarControl {
    const TYPE: ControlType = ControlType::MenuBar;
}

/// Wrapper a MenuItem element as control. The control type of the element must be `UIA_MenuItemControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `ExpandCollapse`, `Invoke`, `SelectionItem`, `Toggle`
#[derive(Debug, Control, ExpandCollapse, Invoke, SelectionItem, Toggle)]
pub struct MenuItemControl {
    control: UIElement
}

impl Control for MenuItemControl {
    const TYPE: ControlType = ControlType::MenuItem;
}

/// Wrapper a Pane element as control. The control type of the element must be `UIA_PaneControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Dock`, `Scroll`, `Transform`, `Window`
#[derive(Debug,	Control, Dock, Scroll, Transform, Window)]
pub struct PaneControl {
    control: UIElement
}

impl Control for PaneControl {
    const TYPE: ControlType = ControlType::Pane;
}

/// Wrapper a ProgressBar element as control. The control type of the element must be `UIA_ProgressBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: RangeValue, Value
#[derive(Debug, Control, RangeValue, Value)]
pub struct ProgressBarControl {
    control: UIElement
}

impl Control for ProgressBarControl {
    const TYPE: ControlType = ControlType::ProgressBar;
}

/// Wrapper a RadioButton element as control. The control type of the element must be `UIA_RadioButtonControlTypeId`.
/// 
/// + Must support: `SelectionItem`
/// + Conditional support: None
#[derive(Debug, Control, SelectionItem)]
pub struct RadioButtonControl {
    control: UIElement
}

impl Control for RadioButtonControl {
    const TYPE: ControlType = ControlType::RadioButton;
}

/// Wrapper a ScrollBar element as control. The control type of the element must be `UIA_ScrollBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `RangeValue`
#[derive(Debug, Control, RangeValue)]
pub struct ScrollBarControl {
    control: UIElement
}

impl Control for ScrollBarControl {
    const TYPE: ControlType = ControlType::ScrollBar;
}

/// Wrapper a SemanticZoom element as control. The control type of the element must be `UIA_SemanticZoomControlTypeId`.
/// 
/// + Must support: `Toggle`
/// + Conditional support: None
#[derive(Debug, Control, Toggle)]
pub struct SemanticZoomControl {
    control: UIElement
}

impl Control for SemanticZoomControl {
    const TYPE: ControlType = ControlType::SemanticZoom;
}

/// Wrapper a Separator element as control. The control type of the element must be `UIA_SeparatorControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: None
#[derive(Debug, Control)]
pub struct SeparatorControl {
    control: UIElement
}

impl Control for SeparatorControl {
    const TYPE: ControlType = ControlType::Separator;
}

/// Wrapper a Slider element as control. The control type of the element must be `UIA_SliderControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `RangeValue`, `Selection`, `Value`
#[derive(Debug, Control, RangeValue, Selection, Value)]
pub struct SliderControl {
    control: UIElement
}

impl Control for SliderControl {
    const TYPE: ControlType = ControlType::Slider;
}

/// Wrapper a Spinner element as control. The control type of the element must be `UIA_SpinnerControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `RangeValue`, `Selection`, `Value`
#[derive(Debug, Control, RangeValue, Selection, Value)]
pub struct SpinnerControl {
    control: UIElement
}

impl Control for SpinnerControl {
    const TYPE: ControlType = ControlType::Spinner;
}

/// Wrapper a SplitButton element as control. The control type of the element must be `UIA_SplitButtonControlTypeId`.
/// 
/// + Must support: `ExpandCollapse`, `Invoke`
/// + Conditional support: None
#[derive(Debug, Control, ExpandCollapse, Invoke)]
pub struct SplitButtonControl {
    control: UIElement
}

impl Control for SplitButtonControl {
    const TYPE: ControlType = ControlType::SplitButton;
}

/// Wrapper a StatusBar element as control. The control type of the element must be `UIA_StatusBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Grid`
#[derive(Debug, Control, Grid)]
pub struct StatusBarControl {
    control: UIElement
}

impl Control for StatusBarControl {
    const TYPE: ControlType = ControlType::StatusBar;
}

/// Wrapper a Tab element as control. The control type of the element must be `UIA_TabControlTypeId`.
/// 
/// + Must support: `Selection`
/// + Conditional support: `Scroll`
#[derive(Debug, Control, Selection, Scroll)]
pub struct TabControl {
    control: UIElement
}

impl Control for TabControl {
    const TYPE: ControlType = ControlType::Tab;
}

/// Wrapper a TabItem element as control. The control type of the element must be `UIA_TabItemControlTypeId`.
/// 
/// + Must support: `SelectionItem`
/// + Conditional support: None
#[derive(Debug, Control, SelectionItem)]
pub struct TabItemControl {
    control: UIElement
}

impl Control for TabItemControl {
    const TYPE: ControlType = ControlType::TabItem;
}

/// Wrapper a Table element as control. The control type of the element must be `UIA_TableControlTypeId`.
/// 
/// + Must support: `Grid`, `GridItem`, `Table`, `TableItem`
/// + Conditional support: None
#[derive(Debug, Control, Grid, GridItem, Table, TableItem)]
pub struct TableControl {
    control: UIElement
}

impl Control for TableControl {
    const TYPE: ControlType = ControlType::Table;
}

/// Wrapper a Text element as control. The control type of the element must be `UIA_TextControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `GridItem`, `TableItem`, `Text`
#[derive(Debug, Control, GridItem, TableItem, Text)]
pub struct TextControl {
    control: UIElement
}

impl Control for TextControl {
    const TYPE: ControlType = ControlType::Text;
}

/// Wrapper a Thumb element as control. The control type of the element must be `UIA_ThumbControlTypeId`.
/// 
/// + Must support: `Transform`
/// + Conditional support: None
#[derive(Debug, Control, Transform)]
pub struct ThumbControl {
    control: UIElement
}

impl Control for ThumbControl {
    const TYPE: ControlType = ControlType::Thumb;
}

/// Wrapper a TitleBar element as control. The control type of the element must be `UIA_TitleBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: None
#[derive(Debug, Control)]
pub struct TitleBarControl {
    control: UIElement
}

impl Control for TitleBarControl {
    const TYPE: ControlType = ControlType::TitleBar;
}

/// Wrapper a ToolBar element as control. The control type of the element must be `UIA_ToolBarControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Dock`, `ExpandCollapse`, `Transform`
#[derive(Debug, Control, Dock, ExpandCollapse, Transform)]
pub struct ToolBarControl {
    control: UIElement
}

impl Control for ToolBarControl {
    const TYPE: ControlType = ControlType::ToolBar;
}

/// Wrapper a ToolTip element as control. The control type of the element must be `UIA_ToolTipControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Text`, `Window`
#[derive(Debug, Control, Text, Window)]
pub struct ToolTipControl {
    control: UIElement
}

impl Control for ToolTipControl {
    const TYPE: ControlType = ControlType::ToolTip;
}

/// Wrapper a Tree element as control. The control type of the element must be `UIA_TreeControlTypeId`.
/// 
/// + Must support: None
/// + Conditional support: `Scroll`, `Selection`
#[derive(Debug,	Control, Scroll, Selection)]
pub struct TreeControl {
    control: UIElement
}

impl Control for TreeControl {
    const TYPE: ControlType = ControlType::Tree;
}

/// Wrapper a TreeItem element as control. The control type of the element must be `UIA_TreeItemControlTypeId`.
/// 
/// + Must support: `ExpandCollapse`
/// + Conditional support: `Invoke`, `ScrollItem`, `SelectionItem`, `Toggle`
#[derive(Debug,	Control, ExpandCollapse, Invoke, ScrollItem, SelectionItem, Toggle)]
pub struct TreeItemControl {
    control: UIElement
}

impl Control for TreeItemControl {
    const TYPE: ControlType = ControlType::TreeItem;
}

/// Wrapper a Window element as control. The control type of the element must be `UIA_WindowControlTypeId`.
/// 
/// + Must support: `Transform`, `Window`
/// + Conditional support: `Dock`
#[derive(Debug, Control, Transform, Window, Dock)]
pub struct WindowControl {
    control: UIElement
}

impl WindowControl {
    /// Brings the thread that created the specified window into the foreground and activates the window. 
    pub fn set_foregrand(&self) -> Result<bool> {
        let hwnd = self.control.get_native_window_handle()?;
        let ret = unsafe {
            SetForegroundWindow(hwnd.into())
        };
        Ok(ret.as_bool())
    }
}

impl Control for WindowControl {
    const TYPE: ControlType = ControlType::Window;
}
