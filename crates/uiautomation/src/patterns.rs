use windows::Win32::Foundation::BOOL;
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::VARIANT;
use windows::Win32::UI::Accessibility::*;
use windows::core::IUnknown;
use windows::core::Interface;

use crate::types::Point;

use super::core::UIElement;
use super::errors::ERR_NOTFOUND;
use super::errors::Error;
use super::errors::Result;
use super::variants::SafeArray;
use super::variants::Variant;

pub trait UIPattern : Sized {
    fn pattern_id() -> i32;
    fn new(pattern: IUnknown) -> Result<Self>;
}

#[derive(Debug, Clone)]
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

impl TryFrom<IUnknown> for UIInvokePattern {
    type Error = Error;

    fn try_from(pattern: IUnknown) -> core::result::Result<Self, Self::Error> {
        let pattern: IUIAutomationInvokePattern = pattern.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationInvokePattern> for UIInvokePattern {
    fn from(pattern: IUIAutomationInvokePattern) -> Self {
        Self {
            pattern
        }
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

#[derive(Debug, Clone)]
pub struct UIAnnotationPattern {
    pattern: IUIAutomationAnnotationPattern
}

impl UIAnnotationPattern {
    pub fn get_type_id(&self) -> Result<i32> {
        let id = unsafe {
            self.pattern.CurrentAnnotationTypeId()?
        };
        Ok(id)
    }

    pub fn get_type_nane(&self) -> Result<String> {
        let name = unsafe {
            self.pattern.CurrentAnnotationTypeName()?
        };
        Ok(name.to_string())
    }

    pub fn get_author(&self) -> Result<String> {
        let author = unsafe {
            self.pattern.CurrentAuthor()?
        };
        Ok(author.to_string())
    }

    pub fn get_datetime(&self) -> Result<String> {
        let datetime = unsafe {
            self.pattern.CurrentDateTime()?
        };
        Ok(datetime.to_string())
    }

    pub fn get_target(&self) -> Result<UIElement> {
        let target = unsafe {
            self.pattern.CurrentTarget()?
        };
        Ok(target.into())
    }
}

impl UIPattern for UIAnnotationPattern {
    fn pattern_id() -> i32 {
        UIA_AnnotationPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIAnnotationPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIAnnotationPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> core::result::Result<Self, Self::Error> {
        let pattern: IUIAutomationAnnotationPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationAnnotationPattern> for UIAnnotationPattern {
    fn from(pattern: IUIAutomationAnnotationPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationAnnotationPattern> for UIAnnotationPattern {
    fn into(self) -> IUIAutomationAnnotationPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationAnnotationPattern> for UIAnnotationPattern {
    fn as_ref(&self) -> &IUIAutomationAnnotationPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UICustomNavigationPattern {
    pattern: IUIAutomationCustomNavigationPattern
}

impl UICustomNavigationPattern {
    pub fn navigate(&self, direction: NavigateDirection) -> Result<UIElement> {
        let element = unsafe {
            self.pattern.Navigate(direction)?
        };
        Ok(element.into())
    }
}

impl UIPattern for UICustomNavigationPattern {
    fn pattern_id() -> i32 {
        UIA_CustomNavigationPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UICustomNavigationPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UICustomNavigationPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> core::result::Result<Self, Self::Error> {
        let pattern: IUIAutomationCustomNavigationPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationCustomNavigationPattern> for UICustomNavigationPattern {
    fn from(pattern: IUIAutomationCustomNavigationPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationCustomNavigationPattern> for UICustomNavigationPattern {
    fn into(self) -> IUIAutomationCustomNavigationPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationCustomNavigationPattern> for UICustomNavigationPattern {
    fn as_ref(&self) -> &IUIAutomationCustomNavigationPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UIDockPattern {
    pattern: IUIAutomationDockPattern
}

impl UIDockPattern {
    pub fn get_dock_position(&self) -> Result<DockPosition> {
        let pos = unsafe {
            self.pattern.CurrentDockPosition()?
        };
        Ok(pos)
    }

    pub fn set_dock_position(&self, position: DockPosition) -> Result<()> {
        unsafe {
            self.pattern.SetDockPosition(position)?
        };
        Ok(())
    }
}

impl UIPattern for UIDockPattern {
    fn pattern_id() -> i32 {
        UIA_DockPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIDockPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIDockPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> core::result::Result<Self, Self::Error> {
        let pattern: IUIAutomationDockPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationDockPattern> for UIDockPattern {
    fn from(pattern: IUIAutomationDockPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationDockPattern> for UIDockPattern {
    fn into(self) -> IUIAutomationDockPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationDockPattern> for UIDockPattern {
    fn as_ref(&self) -> &IUIAutomationDockPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UIDragPattern {
    pattern: IUIAutomationDragPattern
}

impl UIDragPattern {
    pub fn is_grabbed(&self) -> Result<bool> {
        let grabbed = unsafe {
            self.pattern.CurrentIsGrabbed()?
        };
        Ok(grabbed.as_bool())
    }

    pub fn get_drop_effect(&self) -> Result<String> {
        let effect = unsafe {
            self.pattern.CurrentDropEffect()?
        };

        Ok(effect.to_string())
    }

    pub fn get_drop_effects(&self) -> Result<Vec<String>> {
        let effects = unsafe {
            self.pattern.CurrentDropEffects()?
        };

        let effects: SafeArray = effects.into();
        effects.try_into()
    }

    pub fn get_grabbed_items(&self) -> Result<Vec<UIElement>> {
        let elements = unsafe {
            self.pattern.GetCurrentGrabbedItems()?
        };
        let len = unsafe {
            elements.Length()?
        };

        let mut items: Vec<UIElement> = Vec::new();
        for i in 0..len {
            let item = unsafe {
                elements.GetElement(i)?
            };
            let item = UIElement::from(item);
            items.push(item);
        }

        Ok(items)
    }
}

impl UIPattern for UIDragPattern {
    fn pattern_id() -> i32 {
        UIA_DragPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIDragPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIDragPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> core::result::Result<Self, Self::Error> {
        let pattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationDragPattern> for UIDragPattern {
    fn from(pattern: IUIAutomationDragPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationDragPattern> for UIDragPattern {
    fn into(self) -> IUIAutomationDragPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationDragPattern> for UIDragPattern {
    fn as_ref(&self) -> &IUIAutomationDragPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UIDropTargetPattern {
    pattern: IUIAutomationDropTargetPattern
}

impl UIDropTargetPattern {
    pub fn get_drop_target_effect(&self) -> Result<String> {
        let effect = unsafe {
            self.pattern.CurrentDropTargetEffect()?
        };
        Ok(effect.to_string())        
    }

    pub fn get_drop_target_effects(&self) -> Result<Vec<String>> {
        let effects = unsafe {
            self.pattern.CurrentDropTargetEffects()?
        };

        let effects: SafeArray = effects.into();
        effects.try_into()
    }
}

impl UIPattern for UIDropTargetPattern {
    fn pattern_id() -> i32 {
        UIA_DropTargetPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIDropTargetPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIDropTargetPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> core::result::Result<Self, Self::Error> {
        let pattern: IUIAutomationDropTargetPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationDropTargetPattern> for UIDropTargetPattern {
    fn from(pattern: IUIAutomationDropTargetPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationDropTargetPattern> for UIDropTargetPattern {
    fn into(self) -> IUIAutomationDropTargetPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationDropTargetPattern> for UIDropTargetPattern {
    fn as_ref(&self) -> &IUIAutomationDropTargetPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UIExpandCollapsePattern {
    pattern: IUIAutomationExpandCollapsePattern
}

impl UIExpandCollapsePattern {
    pub fn expand(&self) -> Result<()> {
        Ok(unsafe {
            self.pattern.Expand()?
        })
    }

    pub fn collapse(&self) -> Result<()> {
        Ok(unsafe {
            self.pattern.Collapse()?
        })
    }

    pub fn get_state(&self) -> Result<ExpandCollapseState> {
        Ok(unsafe {
            self.pattern.CurrentExpandCollapseState()?
        })
    }
}

impl UIPattern for UIExpandCollapsePattern {
    fn pattern_id() -> i32 {
        UIA_ExpandCollapsePatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIExpandCollapsePattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIExpandCollapsePattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationExpandCollapsePattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationExpandCollapsePattern> for UIExpandCollapsePattern {
    fn from(pattern: IUIAutomationExpandCollapsePattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationExpandCollapsePattern> for UIExpandCollapsePattern {
    fn into(self) -> IUIAutomationExpandCollapsePattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationExpandCollapsePattern> for UIExpandCollapsePattern {
    fn as_ref(&self) -> &IUIAutomationExpandCollapsePattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UIGridPattern {
    pattern: IUIAutomationGridPattern
}

impl UIGridPattern {
    pub fn get_column_count(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentColumnCount()?
        })
    }

    pub fn get_row_count(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentRowCount()?
        })
    }

    pub fn get_item(&self, row: i32, column: i32) -> Result<UIElement> {
        let element = unsafe {
            self.pattern.GetItem(row, column)?
        };
        Ok(UIElement::from(element))
    }
}

impl UIPattern for UIGridPattern {
    fn pattern_id() -> i32 {
        UIA_GridPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIGridPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIGridPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationGridPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationGridPattern> for UIGridPattern {
    fn from(pattern: IUIAutomationGridPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationGridPattern> for UIGridPattern {
    fn into(self) -> IUIAutomationGridPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationGridPattern> for UIGridPattern {
    fn as_ref(&self) -> &IUIAutomationGridPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UIGridItemPattern {
    pattern: IUIAutomationGridItemPattern
}

impl UIGridItemPattern {
    pub fn get_containing_grid(&self) -> Result<UIElement> {
        let grid = unsafe {
            self.pattern.CurrentContainingGrid()?
        };
        Ok(grid.into())
    }

    pub fn get_row(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentRow()?
        })
    }

    pub fn get_column(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentColumn()?
        })
    }

    pub fn get_row_span(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentRowSpan()?
        })
    }

    pub fn get_column_span(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentColumnSpan()?
        })
    }
}

impl UIPattern for UIGridItemPattern {
    fn pattern_id() -> i32 {
        UIA_GridItemPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIGridItemPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIGridItemPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationGridItemPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationGridItemPattern> for UIGridItemPattern {
    fn from(pattern: IUIAutomationGridItemPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationGridItemPattern> for UIGridItemPattern {
    fn into(self) -> IUIAutomationGridItemPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationGridItemPattern> for UIGridItemPattern {
    fn as_ref(&self) -> &IUIAutomationGridItemPattern {
        &self.pattern
    }
}

/// A wrapper for `IUIAutomationItemContainerPattern`.
#[derive(Debug, Clone)]
pub struct UIItemContainerPattern {
    pattern: IUIAutomationItemContainerPattern
}

impl UIItemContainerPattern {
    pub fn find_item_by_property(&self, start_after: UIElement, property_id: i32, value: Variant) -> Result<UIElement> {
        let val: VARIANT = value.into();
        let element = unsafe {
            self.pattern.FindItemByProperty(start_after.as_ref(), property_id, val)?
        };

        Ok(element.into())
    }
}

impl UIPattern for UIItemContainerPattern {
    fn pattern_id() -> i32 {
        UIA_ItemContainerPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIItemContainerPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationItemContainerPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationItemContainerPattern> for UIItemContainerPattern {
    fn from(pattern: IUIAutomationItemContainerPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationItemContainerPattern> for UIItemContainerPattern {
    fn into(self) -> IUIAutomationItemContainerPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationItemContainerPattern> for UIItemContainerPattern {
    fn as_ref(&self) -> &IUIAutomationItemContainerPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UIMultipleViewPattern {
    pattern: IUIAutomationMultipleViewPattern
}

impl UIMultipleViewPattern {
    pub fn get_supported_views(&self) -> Result<Vec<i32>> {
        let views = unsafe {
            self.pattern.GetCurrentSupportedViews()?
        };

        let views: SafeArray = views.into();
        views.try_into()
    }

    pub fn get_view_name(&self, view: i32) -> Result<String> {
        let name = unsafe {
            self.pattern.GetViewName(view)?
        };
        Ok(name.to_string())
    }

    pub fn set_current_view(&self, view: i32) -> Result<()> {
        Ok(unsafe {
            self.pattern.SetCurrentView(view)?
        })
    }

    pub fn get_current_view(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentCurrentView()?
        })
    }
}

impl UIPattern for UIMultipleViewPattern {
    fn pattern_id() -> i32 {
        UIA_MultipleViewPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIMultipleViewPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIMultipleViewPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationMultipleViewPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationMultipleViewPattern> for UIMultipleViewPattern {
    fn from(pattern: IUIAutomationMultipleViewPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationMultipleViewPattern> for UIMultipleViewPattern {
    fn into(self) -> IUIAutomationMultipleViewPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationMultipleViewPattern> for UIMultipleViewPattern {
    fn as_ref(&self) -> &IUIAutomationMultipleViewPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UIRangeValuePattern {
    pattern: IUIAutomationRangeValuePattern
}

impl UIRangeValuePattern {
    pub fn set_value(&self, value: f64) -> Result<()> {
        Ok(unsafe {
            self.pattern.SetValue(value)?
        })
    }

    pub fn get_value(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentValue()?
        })
    }

    pub fn is_readonly(&self) -> Result<bool> {
        let readonly = unsafe {
            self.pattern.CurrentIsReadOnly()?
        };
        Ok(readonly.as_bool())
    }

    pub fn get_maximum(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentMaximum()?
        })
    }

    pub fn get_minimum(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentMinimum()?
        })
    }

    pub fn get_large_change(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentLargeChange()?
        })
    }

    pub fn get_small_change(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentSmallChange()?
        })
    }
}

impl UIPattern for UIRangeValuePattern {
    fn pattern_id() -> i32 {
        UIA_RangeValuePatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIRangeValuePattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIRangeValuePattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationRangeValuePattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationRangeValuePattern> for UIRangeValuePattern {
    fn from(pattern: IUIAutomationRangeValuePattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationRangeValuePattern> for UIRangeValuePattern {
    fn into(self) -> IUIAutomationRangeValuePattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationRangeValuePattern> for UIRangeValuePattern {
    fn as_ref(&self) -> &IUIAutomationRangeValuePattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UIScrollPattern {
    pattern: IUIAutomationScrollPattern
}

impl UIScrollPattern {
    pub fn scroll(&self, horizontal_amount: ScrollAmount, vertical_amount: ScrollAmount) -> Result<()> {
        Ok(unsafe {
            self.pattern.Scroll(horizontal_amount, vertical_amount)?
        })
    }

    pub fn set_scroll_percent(&self, horizontal_percent: f64, vertical_percent: f64) -> Result<()> {
        Ok(unsafe {
            self.pattern.SetScrollPercent(horizontal_percent, vertical_percent)?
        })
    }

    pub fn get_horizontal_scroll_percent(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentHorizontalScrollPercent()?
        })
    }

    pub fn get_vertical_scroll_percent(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentVerticalScrollPercent()?
        })
    }

    pub fn get_horizontal_view_size(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentHorizontalViewSize()?
        })
    }

    pub fn get_vertical_view_size(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentVerticalViewSize()?
        })
    }

    pub fn is_horizontally_scrollable(&self) -> Result<bool> {
        let scrollable = unsafe {
            self.pattern.CurrentHorizontallyScrollable()?
        };
        Ok(scrollable.as_bool())
    }

    pub fn is_vertically_scrollable(&self) -> Result<bool> {
        let scrollable = unsafe {
            self.pattern.CurrentVerticallyScrollable()?
        };
        Ok(scrollable.as_bool())
    }
}

impl UIPattern for UIScrollPattern {
    fn pattern_id() -> i32 {
        UIA_ScrollPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIScrollPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIScrollPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationScrollPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationScrollPattern> for UIScrollPattern {
    fn from(pattern: IUIAutomationScrollPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationScrollPattern> for UIScrollPattern {
    fn into(self) -> IUIAutomationScrollPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationScrollPattern> for UIScrollPattern {
    fn as_ref(&self) -> &IUIAutomationScrollPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UIScrollItemPattern {
    pattern: IUIAutomationScrollItemPattern
}

impl UIScrollItemPattern {
    pub fn scroll_into_view(&self) -> Result<()> {
        Ok(unsafe {
            self.pattern.ScrollIntoView()?
        })
    }
}

impl UIPattern for UIScrollItemPattern {
    fn pattern_id() -> i32 {
        UIA_ScrollItemPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UIScrollItemPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIScrollItemPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationScrollItemPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl Into<IUIAutomationScrollItemPattern> for UIScrollItemPattern {
    fn into(self) -> IUIAutomationScrollItemPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationScrollItemPattern> for UIScrollItemPattern {
    fn as_ref(&self) -> &IUIAutomationScrollItemPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UISelectionPattern {
    pattern: IUIAutomationSelectionPattern
}

impl UISelectionPattern {
    pub fn get_selection(&self) -> Result<Vec<UIElement>> {
        let elem_arr = unsafe {
            self.pattern.GetCurrentSelection()?
        };

        let elements = UIElement::to_elements(elem_arr)?;

        Ok(elements)
    }

    pub fn can_select_multiple(&self) -> Result<bool> {
        let multiple = unsafe {
            self.pattern.CurrentCanSelectMultiple()?
        };
        Ok(multiple.as_bool())
    }

    pub fn is_selection_required(&self) -> Result<bool> {
        let required = unsafe {
            self.pattern.CurrentIsSelectionRequired()?
        };
        Ok(required.as_bool())
    }

    pub fn get_first_selected_item(&self) -> Result<UIElement> {
        let pattern2: IUIAutomationSelectionPattern2 = self.pattern.cast()?;
        let item = unsafe {
            pattern2.CurrentFirstSelectedItem()?
        };
        Ok(item.into())
    }

    pub fn get_last_selected_item(&self) -> Result<UIElement> {
        let pattern2: IUIAutomationSelectionPattern2 = self.pattern.cast()?;
        let item = unsafe {
            pattern2.CurrentLastSelectedItem()?
        };
        Ok(item.into())
    }

    pub fn get_current_selected_item(&self) -> Result<UIElement> {
        let pattern2: IUIAutomationSelectionPattern2 = self.pattern.cast()?;
        let item = unsafe {
            pattern2.CurrentCurrentSelectedItem()?
        };
        Ok(item.into())
    }

    pub fn get_item_count(&self) -> Result<i32> {
        let pattern2: IUIAutomationSelectionPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CurrentItemCount()?
        })
    }
}

impl UIPattern for UISelectionPattern {
    fn pattern_id() -> i32 {
        UIA_SelectionPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UISelectionPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UISelectionPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationSelectionPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationSelectionPattern> for UISelectionPattern {
    fn from(pattern: IUIAutomationSelectionPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationSelectionPattern> for UISelectionPattern {
    fn into(self) -> IUIAutomationSelectionPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationSelectionPattern> for UISelectionPattern {
    fn as_ref(&self) -> &IUIAutomationSelectionPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UISelectionItemPattern {
    pattern: IUIAutomationSelectionItemPattern
}

impl UISelectionItemPattern {
    pub fn select(&self) -> Result<()> {
        Ok(unsafe {
            self.pattern.Select()?
        })
    }

    pub fn add_to_selection(&self) -> Result<()> {
        Ok(unsafe {
            self.pattern.AddToSelection()?
        })
    }

    pub fn remove_from_selection(&self) -> Result<()> {
        Ok(unsafe {
            self.pattern.RemoveFromSelection()?
        })
    }

    pub fn is_selected(&self) -> Result<bool> {
        let selected = unsafe {
            self.pattern.CurrentIsSelected()?
        };
        Ok(selected.as_bool())
    }

    pub fn get_selection_container(&self) -> Result<UIElement> {
        let container = unsafe {
            self.pattern.CurrentSelectionContainer()?
        };
        Ok(container.into())
    }
}

impl UIPattern for UISelectionItemPattern {
    fn pattern_id() -> i32 {
        UIA_SelectionItemPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        UISelectionItemPattern::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UISelectionItemPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationSelectionItemPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationSelectionItemPattern> for UISelectionItemPattern {
    fn from(pattern: IUIAutomationSelectionItemPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationSelectionItemPattern> for UISelectionItemPattern {
    fn into(self) -> IUIAutomationSelectionItemPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationSelectionItemPattern> for UISelectionItemPattern {
    fn as_ref(&self) -> &IUIAutomationSelectionItemPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UISpreadsheetPattern {
    pattern: IUIAutomationSpreadsheetPattern
}

impl UISpreadsheetPattern {
    pub fn get_item_by_name(&self, name: &str) -> Result<UIElement> {
        let name = BSTR::from(name);
        let item = unsafe {
            self.pattern.GetItemByName(name)?
        };
        Ok(item.into())
    }
}

impl UIPattern for UISpreadsheetPattern {
    fn pattern_id() -> i32 {
        UIA_SpreadsheetPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UISpreadsheetPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationSpreadsheetPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationSpreadsheetPattern> for UISpreadsheetPattern {
    fn from(pattern: IUIAutomationSpreadsheetPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationSpreadsheetPattern> for UISpreadsheetPattern {
    fn into(self) -> IUIAutomationSpreadsheetPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationSpreadsheetPattern> for UISpreadsheetPattern {
    fn as_ref(&self) -> &IUIAutomationSpreadsheetPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UISpreadsheetItemPattern {
    pattern: IUIAutomationSpreadsheetItemPattern
}

impl UISpreadsheetItemPattern {
    pub fn get_formula(&self) -> Result<String> {
        let formula = unsafe {
            self.pattern.CurrentFormula()?
        };
        Ok(formula.to_string())
    }

    pub fn get_annotation_objects(&self) -> Result<Vec<UIElement>> {
        let elem_arr = unsafe {
            self.pattern.GetCurrentAnnotationObjects()?
        };
        let elements = UIElement::to_elements(elem_arr)?;
        Ok(elements)
    }

    pub fn get_annotation_types(&self) -> Result<Vec<i32>> {
        let types = unsafe {
            self.pattern.GetCurrentAnnotationTypes()?
        };

        let types: SafeArray = types.into();
        types.try_into()
    }
}

impl UIPattern for UISpreadsheetItemPattern {
    fn pattern_id() -> i32 {
        UIA_SpreadsheetItemPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UISpreadsheetItemPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationSpreadsheetItemPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationSpreadsheetItemPattern> for UISpreadsheetItemPattern {
    fn from(pattern: IUIAutomationSpreadsheetItemPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationSpreadsheetItemPattern> for UISpreadsheetItemPattern {
    fn into(self) -> IUIAutomationSpreadsheetItemPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationSpreadsheetItemPattern> for UISpreadsheetItemPattern {
    fn as_ref(&self) -> &IUIAutomationSpreadsheetItemPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UIStylesPattern {
    pattern: IUIAutomationStylesPattern
}

impl UIStylesPattern {
    pub fn get_style_id(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentStyleId()?
        })
    }

    pub fn get_style_name(&self) -> Result<String> {
        let name = unsafe {
            self.pattern.CurrentStyleName()?
        };
        Ok(name.to_string())
    }

    pub fn get_fill_color(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentFillColor()?
        })
    }

    pub fn get_fill_pattern_style(&self) -> Result<String> {
        let style = unsafe {
            self.pattern.CurrentFillPatternStyle()?
        };
        Ok(style.to_string())
    }

    pub fn get_fill_pattern_color(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentFillPatternColor()?
        })
    }

    pub fn get_shape(&self) -> Result<String> {
        let shape = unsafe {
            self.pattern.CurrentShape()?
        };
        Ok(shape.to_string())
    }

    pub fn get_extended_properties(&self) -> Result<String> {
        let properties = unsafe {
            self.pattern.CurrentExtendedProperties()?
        };
        Ok(properties.to_string())
    }
}

impl UIPattern for UIStylesPattern {
    fn pattern_id() -> i32 {
        UIA_StylesPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIStylesPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationStylesPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationStylesPattern> for UIStylesPattern {
    fn from(pattern: IUIAutomationStylesPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationStylesPattern> for UIStylesPattern {
    fn into(self) -> IUIAutomationStylesPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationStylesPattern> for UIStylesPattern {
    fn as_ref(&self) -> &IUIAutomationStylesPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UISynchronizedInputPattern {
    pattern: IUIAutomationSynchronizedInputPattern
}

impl UISynchronizedInputPattern {
    pub fn start_listening(&self, input_type: SynchronizedInputType) -> Result<()> {
        Ok(unsafe {
            self.pattern.StartListening(input_type)?
        })
    }

    pub fn cancel(&self) -> Result<()> {
        Ok(unsafe {
            self.pattern.Cancel()?
        })
    }
}

impl UIPattern for UISynchronizedInputPattern {
    fn pattern_id() -> i32 {
        UIA_SynchronizedInputPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UISynchronizedInputPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationSynchronizedInputPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationSynchronizedInputPattern> for UISynchronizedInputPattern {
    fn from(pattern: IUIAutomationSynchronizedInputPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationSynchronizedInputPattern> for UISynchronizedInputPattern {
    fn into(self) -> IUIAutomationSynchronizedInputPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationSynchronizedInputPattern> for UISynchronizedInputPattern {
    fn as_ref(&self) -> &IUIAutomationSynchronizedInputPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UITablePattern {
    pattern: IUIAutomationTablePattern
}

impl UITablePattern {
    pub fn get_row_headers(&self) -> Result<Vec<UIElement>> {
        let headers = unsafe {
            self.pattern.GetCurrentRowHeaders()?
        };

        UIElement::to_elements(headers)
    }

    pub fn get_column_headers(&self) -> Result<Vec<UIElement>> {
        let headers = unsafe {
            self.pattern.GetCurrentColumnHeaders()?
        };

        UIElement::to_elements(headers)
    }

    pub fn get_row_or_column_major(&self) -> Result<RowOrColumnMajor> {
        Ok(unsafe {
            self.pattern.CurrentRowOrColumnMajor()?
        })
    }
}

impl UIPattern for UITablePattern {
    fn pattern_id() -> i32 {
        UIA_TablePatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UITablePattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationTablePattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationTablePattern> for UITablePattern {
    fn from(pattern: IUIAutomationTablePattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationTablePattern> for UITablePattern {
    fn into(self) -> IUIAutomationTablePattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationTablePattern> for UITablePattern {
    fn as_ref(&self) -> &IUIAutomationTablePattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UITableItemPattern {
    pattern: IUIAutomationTableItemPattern
}

impl UITableItemPattern {
    pub fn get_row_header_items(&self) -> Result<Vec<UIElement>> {
        let items = unsafe {
            self.pattern.GetCurrentRowHeaderItems()?
        };

        UIElement::to_elements(items)
    }

    pub fn get_column_header_items(&self) -> Result<Vec<UIElement>> {
        let items = unsafe {
            self.pattern.GetCurrentColumnHeaderItems()?
        };

        UIElement::to_elements(items)
    }
}

impl UIPattern for UITableItemPattern {
    fn pattern_id() -> i32 {
        UIA_TableItemPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UITableItemPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationTableItemPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationTableItemPattern> for UITableItemPattern {
    fn from(pattern: IUIAutomationTableItemPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationTableItemPattern> for UITableItemPattern {
    fn into(self) -> IUIAutomationTableItemPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationTableItemPattern> for UITableItemPattern {
    fn as_ref(&self) -> &IUIAutomationTableItemPattern {
        &self.pattern
    }
}

#[derive(Debug, Clone)]
pub struct UITextChildPattern {
    pattern: IUIAutomationTextChildPattern
}

impl UITextChildPattern {
    pub fn get_text_container(&self) -> Result<UIElement> {
        let container = unsafe {
            self.pattern.TextContainer()?
        };

        Ok(container.into())
    }

    pub fn get_text_range(&self) -> Result<UITextRange> {
        let range = unsafe {
            self.pattern.TextRange()?
        };

        Ok(range.into())
    }
}

impl UIPattern for UITextChildPattern {
    fn pattern_id() -> i32 {
        UIA_TextChildPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UITextChildPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationTextChildPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationTextChildPattern> for UITextChildPattern {
    fn from(pattern: IUIAutomationTextChildPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationTextChildPattern> for UITextChildPattern {
    fn into(self) -> IUIAutomationTextChildPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationTextChildPattern> for UITextChildPattern {
    fn as_ref(&self) -> &IUIAutomationTextChildPattern {
        &self.pattern
    }
}

/// A Wrapper for `IUIAutomationTextPattern` and `IUIAutomationTextPattern2`
#[derive(Debug, Clone)]
pub struct UITextPattern {
    pattern: IUIAutomationTextPattern
}

impl UITextPattern {
    pub fn get_ragne_from_point(&self, pt: Point) -> Result<UITextRange> {
        let range = unsafe {
            self.pattern.RangeFromPoint(pt)?
        };

        Ok(range.into())
    }

    pub fn get_range_from_child(&self, child: &UIElement) -> Result<UITextRange> {
        let range = unsafe {
            self.pattern.RangeFromChild(child.as_ref().clone())?
        };

        Ok(range.into())
    }

    pub fn get_selection(&self) -> Result<Vec<UITextRange>> {
        let ranges = unsafe {
            self.pattern.GetSelection()?
        };

        UITextRange::to_ranges(ranges)
    }

    pub fn get_visible_ranges(&self) -> Result<Vec<UITextRange>> {
        let ranges = unsafe {
            self.pattern.GetVisibleRanges()?
        };

        UITextRange::to_ranges(ranges)
    }

    pub fn get_document_range(&self) -> Result<UITextRange> {
        let range = unsafe {
            self.pattern.DocumentRange()?
        };

        Ok(range.into())
    }

    pub fn get_supported_text_selection(&self) -> Result<SupportedTextSelection> {
        Ok(unsafe {
            self.pattern.SupportedTextSelection()?
        })
    }

    pub fn get_range_from_annotation(&self, annotation: &UIElement) -> Result<UITextRange> {
        let pattern2: IUIAutomationTextPattern2 = self.pattern.cast()?;
        let range = unsafe {
            pattern2.RangeFromAnnotation(annotation.as_ref())?
        };
        Ok(range.into())
    }

    pub fn get_caret_range(&self) -> Result<(bool, UITextRange)> {
        let pattern2: IUIAutomationTextPattern2 = self.pattern.cast()?;
        let mut active = BOOL::default();
        let mut range = Option::None;
        unsafe {
            pattern2.GetCaretRange(&mut active, &mut range)?
        };

        if range.is_some() {
            let range = range.unwrap();
            Ok((active.as_bool(), range.into()))
        } else {
            Err(Error::new(ERR_NOTFOUND, "Range Not Found"))
        }
    }
}

impl UIPattern for UITextPattern {
    fn pattern_id() -> i32 {
        UIA_TextPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UITextPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationTextPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationTextPattern> for UITextPattern {
    fn from(pattern: IUIAutomationTextPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationTextPattern> for UITextPattern {
    fn into(self) -> IUIAutomationTextPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationTextPattern> for UITextPattern {
    fn as_ref(&self) -> &IUIAutomationTextPattern {
        &self.pattern
    }
}

/// A Wrapper for `IUIAutomationTextEditPattern`.
/// 
/// This type inherits from `UITextPattern`.
/// 
#[derive(Debug, Clone)]
pub struct UITextEditPattern {
    text: UITextPattern,
    pattern: IUIAutomationTextEditPattern
}

impl UITextEditPattern {
    pub fn get_active_composition(&self) -> Result<UITextRange> {
        let range = unsafe {
            self.pattern.GetActiveComposition()?
        };
        Ok(range.into())
    }

    pub fn get_conversion_target(&self) -> Result<UITextRange> {
        let range = unsafe {
            self.pattern.GetConversionTarget()?
        };
        Ok(range.into())
    }
}

impl UIPattern for UITextEditPattern {
    fn pattern_id() -> i32 {
        UIA_TextEditPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UITextEditPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationTextEditPattern = value.cast()?;
        let text = UITextPattern::try_from(value)?;
        Ok(Self {
            text,
            pattern
        })
    }
}

impl From<IUIAutomationTextEditPattern> for UITextEditPattern {
    fn from(pattern: IUIAutomationTextEditPattern) -> Self {
        let text: IUIAutomationTextPattern = pattern.cast().unwrap();
        Self {
            text: text.into(),
            pattern
        }
    }
}

impl Into<IUIAutomationTextEditPattern> for UITextEditPattern {
    fn into(self) -> IUIAutomationTextEditPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationTextEditPattern> for UITextEditPattern {
    fn as_ref(&self) -> &IUIAutomationTextEditPattern {
        &self.pattern
    }
}

impl AsRef<UITextPattern> for UITextEditPattern {
    fn as_ref(&self) -> &UITextPattern {
        &self.text
    }
}
/// A wrapper for `IUIAutomationTextRange`, `IUIAutomationTextRange2` and `IUIAutomationTextRange3`.
#[derive(Debug, Clone)]
pub struct UITextRange {
    range: IUIAutomationTextRange
}

impl UITextRange {
    pub fn compare(&self, range: &UITextRange) -> Result<bool> {
        let ret = unsafe {
            self.range.Compare(range.as_ref())?
        };
        Ok(ret.as_bool())
    }

    pub fn compare_endpoints(&self, src_endpoint: TextPatternRangeEndpoint, range: &UITextRange, target_endpoint: TextPatternRangeEndpoint) -> Result<i32> {
       Ok(unsafe {
           self.range.CompareEndpoints(src_endpoint, range.as_ref(), target_endpoint)?
       }) 
    }

    pub fn expand_to_enclosing_unit(&self, text_unit: TextUnit) -> Result<()> {
        Ok(unsafe {
            self.range.ExpandToEnclosingUnit(text_unit)?
        })
    }

    pub fn find_attribute(&self, attr: i32, value: Variant, backward: bool) -> Result<UITextRange> {
        let range = unsafe {
            self.range.FindAttribute(attr, value.as_ref(), backward)?
        };
        Ok(range.into())
    }

    pub fn find_text(&self, text: &str, backward: bool, ignorecase: bool) -> Result<UITextRange> {
        let range = unsafe {
            self.range.FindText(BSTR::from(text), backward, ignorecase)?
        };
        Ok(range.into())
    }

    pub fn get_attribute_value(&self, attr: i32) -> Result<Variant> {
        let value = unsafe {
            self.range.GetAttributeValue(attr)?
        };
        Ok(value.into())
    }

    pub fn get_enclosing_element(&self) -> Result<UIElement> {
        let element = unsafe {
            self.range.GetEnclosingElement()?
        };
        Ok(element.into())
    }

    pub fn get_text(&self, max_length: i32) -> Result<String> {
        let text = unsafe {
            self.range.GetText(max_length)?
        };
        Ok(text.to_string())
    }

    pub fn move_text(&self, unit: TextUnit, count: i32) -> Result<i32> {
        Ok(unsafe {
            self.range.Move(unit, count)?
        })
    }

    pub fn move_endpoint_by_unit(&self, endpoint: TextPatternRangeEndpoint, unit: TextUnit, count: i32) -> Result<i32> {
        Ok(unsafe {
            self.range.MoveEndpointByUnit(endpoint, unit, count)?
        })
    }

    pub fn move_endpoint_by_range(&self, src_endpoint: TextPatternRangeEndpoint, range: &UITextRange, target_endpoint: TextPatternRangeEndpoint) -> Result<()> {
        Ok(unsafe {
            self.range.MoveEndpointByRange(src_endpoint, range.as_ref(), target_endpoint)?
        })
    }

    pub fn select(&self) -> Result<()> {
        Ok(unsafe {
            self.range.Select()?
        })
    }

    pub fn add_to_selection(&self) -> Result<()> {
        Ok(unsafe {
            self.range.AddToSelection()?
        })
    }

    pub fn remove_from_selection(&self) -> Result<()> {
        Ok(unsafe {
            self.range.RemoveFromSelection()?
        })
    }

    pub fn scroll_into_view(&self, align_to_top: bool) -> Result<()> {
        Ok(unsafe {
            self.range.ScrollIntoView(align_to_top)?
        })
    }

    pub fn get_children(&self) -> Result<Vec<UIElement>> {
        let children = unsafe {
            self.range.GetChildren()?
        };

        UIElement::to_elements(children)
    }

    pub fn show_context_menu(&self) -> Result<()> {
        let range2: IUIAutomationTextRange2 = self.range.cast()?;
        Ok(unsafe {
            range2.ShowContextMenu()?
        })
    }

    /// Convert `IUIAutomationTextRangeArray` to `Vec<UITextRange>`.
    pub(crate) fn to_ranges(ranges: IUIAutomationTextRangeArray) -> Result<Vec<UITextRange>> {
        let mut arr: Vec<UITextRange> = Vec::new();

        unsafe {
            for i in 0..ranges.Length()? {
                let range = ranges.GetElement(i)?;
                arr.push(range.into());
            }
        }

        Ok(arr)
    }
}

impl From<IUIAutomationTextRange> for UITextRange {
    fn from(range: IUIAutomationTextRange) -> Self {
        Self {
            range
        }
    }
}

impl Into<IUIAutomationTextRange> for UITextRange {
    fn into(self) -> IUIAutomationTextRange {
        self.range
    }
}

impl AsRef<IUIAutomationTextRange> for UITextRange {
    fn as_ref(&self) -> &IUIAutomationTextRange {
        &self.range
    }
}

/// A wrapper for `IUIAutomationTogglePattern`
#[derive(Debug, Clone)]
pub struct UITogglePattern {
    pattern: IUIAutomationTogglePattern
}

impl UITogglePattern {
    pub fn get_toggle_state(&self) -> Result<ToggleState> {
        Ok(unsafe {
            self.pattern.CurrentToggleState()?
        })
    }

    pub fn toggle(&self) -> Result<()> {
        Ok(unsafe {
            self.pattern.Toggle()?
        })
    }
}

impl UIPattern for UITogglePattern {
    fn pattern_id() -> i32 {
        UIA_TogglePatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UITogglePattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationTogglePattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationTogglePattern> for UITogglePattern {
    fn from(pattern: IUIAutomationTogglePattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationTogglePattern> for UITogglePattern {
    fn into(self) -> IUIAutomationTogglePattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationTogglePattern> for UITogglePattern {
    fn as_ref(&self) -> &IUIAutomationTogglePattern {
        &self.pattern
    }
}

/// A wrapper for `IUIAutomationTransformPattern` and `IUIAutomationTransformPattern2`
#[derive(Debug, Clone)]
pub struct UITransformPattern {
    pattern: IUIAutomationTransformPattern
}

impl UITransformPattern {
    pub fn can_move(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CurrentCanMove()?
        };
        Ok(ret.as_bool())
    }

    pub fn move_to(&self, x: f64, y: f64) -> Result<()> {
        Ok(unsafe {
            self.pattern.Move(x, y)?
        })
    }

    pub fn can_resize(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CurrentCanResize()?
        };
        Ok(ret.as_bool())        
    }

    pub fn resize(&self, width: f64, height: f64) -> Result<()> {
        Ok(unsafe {
            self.pattern.Resize(width, height)?
        })
    }

    pub fn can_rotate(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CurrentCanRotate()?
        };
        Ok(ret.as_bool())
    }

    pub fn rotate(&self, degrees: f64) -> Result<()> {
        Ok(unsafe {
            self.pattern.Rotate(degrees)?
        })
    }

    pub fn can_zoom(&self) -> Result<bool> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        let zoomable = unsafe {
            pattern2.CurrentCanZoom()?
        };
        Ok(zoomable.as_bool())
    }

    pub fn get_zoom_level(&self) -> Result<f64> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CurrentZoomLevel()?
        })
    }

    pub fn get_zoom_minimum(&self) -> Result<f64> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CurrentZoomMinimum()?
        })
    }

    pub fn get_zoom_maximum(&self) -> Result<f64> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CurrentZoomMaximum()?
        })
    }

    pub fn zoom(&self, zoom_value: f64) -> Result<()> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.Zoom(zoom_value)?
        })
    }

    pub fn zoom_by_unit(&self, zoom_unit: ZoomUnit) -> Result<()> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.ZoomByUnit(zoom_unit)?
        })
    }
}

impl UIPattern for UITransformPattern {
    fn pattern_id() -> i32 {
        UIA_TransformPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UITransformPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationTransformPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationTransformPattern> for UITransformPattern {
    fn from(pattern: IUIAutomationTransformPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationTransformPattern> for UITransformPattern {
    fn into(self) -> IUIAutomationTransformPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationTransformPattern> for UITransformPattern {
    fn as_ref(&self) -> &IUIAutomationTransformPattern {
        &self.pattern
    }
}

/// A wrapper for `IUIAutomationValuePattern`.
#[derive(Debug, Clone)]
pub struct UIValuePattern {
    pattern: IUIAutomationValuePattern
}

impl UIValuePattern {
    pub fn set_value(&self, value: &str) -> Result<()> {
        Ok(unsafe {
            self.pattern.SetValue(BSTR::from(value))?
        })
    }

    pub fn get_value(&self) -> Result<String> {
        let value = unsafe {
            self.pattern.CurrentValue()?
        };
        Ok(value.to_string())
    }

    pub fn is_readonly(&self) -> Result<bool> {
        let readonly = unsafe {
            self.pattern.CurrentIsReadOnly()?
        };
        Ok(readonly.as_bool())
    }
}

impl UIPattern for UIValuePattern {
    fn pattern_id() -> i32 {
        UIA_ValuePatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIValuePattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationValuePattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationValuePattern> for UIValuePattern {
    fn from(pattern: IUIAutomationValuePattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationValuePattern> for UIValuePattern {
    fn into(self) -> IUIAutomationValuePattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationValuePattern> for UIValuePattern {
    fn as_ref(&self) -> &IUIAutomationValuePattern {
        &self.pattern
    }
}

/// A wrapper for `IUIAutomationVirtualizedItemPattern`.
#[derive(Debug, Clone)]
pub struct UIVirtualizedItemPattern {
    pattern: IUIAutomationVirtualizedItemPattern
}

impl UIVirtualizedItemPattern {
    pub fn realize(&self) -> Result<()> {
        Ok(unsafe {
            self.pattern.Realize()?
        })
    }
}

impl UIPattern for UIVirtualizedItemPattern {
    fn pattern_id() -> i32 {
        UIA_VirtualizedItemPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIVirtualizedItemPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationVirtualizedItemPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationVirtualizedItemPattern> for UIVirtualizedItemPattern {
    fn from(pattern: IUIAutomationVirtualizedItemPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationVirtualizedItemPattern> for UIVirtualizedItemPattern {
    fn into(self) -> IUIAutomationVirtualizedItemPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationVirtualizedItemPattern> for UIVirtualizedItemPattern {
    fn as_ref(&self) -> &IUIAutomationVirtualizedItemPattern {
        &self.pattern
    }
}

/// A wrapper for `IUIAutomationWindowPattern`.
#[derive(Debug, Clone)]
pub struct UIWindowPattern {
    pattern: IUIAutomationWindowPattern
}

impl UIWindowPattern {
    pub fn close(&self) -> Result<()> {
        Ok(unsafe {
            self.pattern.Close()?
        })
    }

    pub fn wait_for_input_idle(&self, milliseconds: i32) -> Result<bool> {
        let ret = unsafe {
            self.pattern.WaitForInputIdle(milliseconds)?
        };
        Ok(ret.as_bool())
    }

    pub fn get_window_visual_state(&self) -> Result<WindowVisualState> {
        Ok(unsafe {
            self.pattern.CurrentWindowVisualState()?
        })
    }

    pub fn set_window_visual_state(&self, state: WindowVisualState) -> Result<()> {
        Ok(unsafe {
            self.pattern.SetWindowVisualState(state)?
        })
    }

    pub fn can_maximize(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CurrentCanMaximize()?
        };
        Ok(ret.as_bool())
    }

    pub fn can_minimize(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CurrentCanMinimize()?
        };
        Ok(ret.as_bool())
    }

    pub fn is_modal(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CurrentIsModal()?
        };
        Ok(ret.as_bool())
    }

    pub fn is_topmost(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CurrentIsTopmost()?
        };
        Ok(ret.as_bool())
    }

    pub fn get_window_interaction_state(&self) -> Result<WindowInteractionState> {
        Ok(unsafe {
            self.pattern.CurrentWindowInteractionState()?
        })
    }
}

impl UIPattern for UIWindowPattern {
    fn pattern_id() -> i32 {
        UIA_WindowPatternId
    }

    fn new(pattern: IUnknown) -> Result<Self> {
        Self::try_from(pattern)
    }
}

impl TryFrom<IUnknown> for UIWindowPattern {
    type Error = Error;

    fn try_from(value: IUnknown) -> Result<Self> {
        let pattern: IUIAutomationWindowPattern = value.cast()?;
        Ok(Self {
            pattern
        })
    }
}

impl From<IUIAutomationWindowPattern> for UIWindowPattern {
    fn from(pattern: IUIAutomationWindowPattern) -> Self {
        Self {
            pattern
        }
    }
}

impl Into<IUIAutomationWindowPattern> for UIWindowPattern {
    fn into(self) -> IUIAutomationWindowPattern {
        self.pattern
    }
}

impl AsRef<IUIAutomationWindowPattern> for UIWindowPattern {
    fn as_ref(&self) -> &IUIAutomationWindowPattern {
        &self.pattern
    }
}