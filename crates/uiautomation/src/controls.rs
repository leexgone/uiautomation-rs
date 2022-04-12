use uiautomation_derive::Invoke;
use uiautomation_derive::SelectionItem;
use windows::Win32::UI::Accessibility::UIA_ButtonControlTypeId;
use windows::Win32::UI::Accessibility::UIA_ListControlTypeId;
use windows::Win32::UI::Accessibility::UIA_ListItemControlTypeId;

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
use crate::variants::Variant;

/// Wrapper an button element as a control.
#[derive(Invoke)]
pub struct ButtonControl {
    control: UIElement
}

// impl ButtonControl {
//     /// Perform a click event on this control.
//     pub fn click(&self) -> Result<()> {
//         let pattern: UIInvokePattern = self.as_ref().get_pattern()?;
//         pattern.invoke()
//     }
// }

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

impl AsRef<UIElement> for ButtonControl {
    fn as_ref(&self) -> &UIElement {
        &self.control
    }
}

/// Wrapper a list element as a control. The control type of the element must be `UIA_ListControlTypeId`.
pub struct ListControl {
    control: UIElement
}

impl ListControl {
    /// Find contained item by property value.
    /// 
    /// Search `UIElement` from the `start_ater` item, return the first item which property is specified value.
    pub fn find_item_by_property(&self, start_after: UIElement, property_id: i32, value: Variant) -> Result<UIElement> {
        let pattern: UIItemContainerPattern = self.control.get_pattern()?;
        pattern.find_item_by_property(start_after, property_id, value)
    }

    /// Get supported view ids.
    pub fn get_supported_views(&self) -> Result<Vec<i32>> {
        let pattern: UIMultipleViewPattern = self.control.get_pattern()?;
        pattern.get_supported_views()
    }

    /// Return the view name.
    /// 
    /// The `view` parameter is the id of the view. You can get all view ids by `get_supported_views()` function.
    pub fn get_view_name(&self, view: i32) -> Result<String> {
        let pattern: UIMultipleViewPattern = self.control.get_pattern()?;
        pattern.get_view_name(view)
    }

    /// Get the current view id.
    pub fn get_current_view(&self) -> Result<i32> {
        let pattern: UIMultipleViewPattern = self.control.get_pattern()?;
        pattern.get_current_view()
    }

    /// Set the current view by id.
    pub fn set_current_view(&self, view: i32) -> Result<()> {
        let pattern: UIMultipleViewPattern = self.control.get_pattern()?;
        pattern.set_current_view(view)
    }
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

/// Wrapper a listitem element as a control. The control type of the element must be `UIA_ListItemControlTypeId`.
#[derive(Invoke, SelectionItem)]
pub struct ListItemControl {
    control: UIElement
}

impl ListItemControl {
    // /// Determines whether this control is clickable.
    // pub fn is_clickable(&self) -> bool {
    //     let element: &IUIAutomationElement = self.control.as_ref();
    //     let pattern = unsafe { 
    //         element.GetCurrentPattern(UIA_InvokePatternId)
    //     };
    //     pattern.is_ok()
    // }

    // /// Perform a click event on this control.
    // pub fn click(&self) -> Result<()> {
    //     let pattern: UIInvokePattern = self.control.get_pattern()?;
    //     pattern.invoke()
    // }

    /// Scroll this item into view.
    pub fn scroll_into_view(&self) -> Result<()> {
        let pattern: UIScrollItemPattern = self.control.get_pattern()?;
        pattern.scroll_into_view()
    }

    // /// Determines whether this control is slectable.
    // pub fn is_selectable(&self) -> bool {
    //     let element: &IUIAutomationElement = self.control.as_ref();
    //     let pattern = unsafe {
    //         element.GetCurrentPattern(UIA_SelectionItemPatternId)
    //     };
    //     pattern.is_ok()
    // }

    // /// Select current item.
    // pub fn select(&self) -> Result<()> {
    //     let pattern: UISelectionItemPattern = self.as_ref().get_pattern()?;
    //     pattern.select()
    // }

    // /// Add current item to selection.
    // pub fn add_to_selection(&self) -> Result<()> {
    //     let pattern: UISelectionItemPattern = self.as_ref().get_pattern()?;
    //     pattern.add_to_selection()
    // }

    // /// Remove current item from selection.
    // pub fn remove_from_selection(&self) -> Result<()> {
    //     let pattern: UISelectionItemPattern = self.as_ref().get_pattern()?;
    //     pattern.remove_from_selection()
    // }

    // /// Determines whether this item is selected.
    // pub fn is_selected(&self) -> Result<bool> {
    //     let pattern: UISelectionItemPattern = self.as_ref().get_pattern()?;
    //     pattern.is_selected()
    // }
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