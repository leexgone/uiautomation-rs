use windows::Win32::UI::Accessibility::DockPosition;
use windows::Win32::UI::Accessibility::ExpandCollapseState;
use windows::Win32::UI::Accessibility::IUIAutomationAnnotationPattern;
use windows::Win32::UI::Accessibility::IUIAutomationCustomNavigationPattern;
use windows::Win32::UI::Accessibility::IUIAutomationDockPattern;
use windows::Win32::UI::Accessibility::IUIAutomationDragPattern;
use windows::Win32::UI::Accessibility::IUIAutomationDropTargetPattern;
use windows::Win32::UI::Accessibility::IUIAutomationExpandCollapsePattern;
use windows::Win32::UI::Accessibility::IUIAutomationGridItemPattern;
use windows::Win32::UI::Accessibility::IUIAutomationGridPattern;
use windows::Win32::UI::Accessibility::IUIAutomationInvokePattern;
use windows::Win32::UI::Accessibility::IUIAutomationMultipleViewPattern;
use windows::Win32::UI::Accessibility::IUIAutomationRangeValuePattern;
use windows::Win32::UI::Accessibility::IUIAutomationScrollItemPattern;
use windows::Win32::UI::Accessibility::IUIAutomationScrollPattern;
use windows::Win32::UI::Accessibility::IUIAutomationSelectionItemPattern;
use windows::Win32::UI::Accessibility::IUIAutomationSelectionPattern;
use windows::Win32::UI::Accessibility::IUIAutomationSelectionPattern2;
use windows::Win32::UI::Accessibility::NavigateDirection;
use windows::Win32::UI::Accessibility::ScrollAmount;
use windows::Win32::UI::Accessibility::UIA_AnnotationPatternId;
use windows::Win32::UI::Accessibility::UIA_CustomNavigationPatternId;
use windows::Win32::UI::Accessibility::UIA_DockPatternId;
use windows::Win32::UI::Accessibility::UIA_DragPatternId;
use windows::Win32::UI::Accessibility::UIA_DropTargetPatternId;
use windows::Win32::UI::Accessibility::UIA_ExpandCollapsePatternId;
use windows::Win32::UI::Accessibility::UIA_GridItemPatternId;
use windows::Win32::UI::Accessibility::UIA_GridPatternId;
use windows::Win32::UI::Accessibility::UIA_InvokePatternId;
use windows::Win32::UI::Accessibility::UIA_MultipleViewPatternId;
use windows::Win32::UI::Accessibility::UIA_RangeValuePatternId;
use windows::Win32::UI::Accessibility::UIA_ScrollItemPatternId;
use windows::Win32::UI::Accessibility::UIA_ScrollPatternId;
use windows::Win32::UI::Accessibility::UIA_SelectionItemPatternId;
use windows::Win32::UI::Accessibility::UIA_SelectionPatternId;
use windows::core::IUnknown;
use windows::core::Interface;

use crate::core::UIElement;
use crate::core::to_elements;
use crate::errors::Error;
use crate::errors::Result;

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

    pub fn set_doc_position(&self, position: DockPosition) -> Result<()> {
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

#[derive(Debug, Clone)]
pub struct UIMultipleViewPattern {
    pattern: IUIAutomationMultipleViewPattern
}

impl UIMultipleViewPattern {
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

    pub fn get_supported_views(&self) -> Result<Vec<i32>> {
        todo!("get supported views")
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

        let elements = to_elements(elem_arr)?;

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