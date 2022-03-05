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
use windows::Win32::UI::Accessibility::NavigateDirection;
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
use windows::core::IUnknown;
use windows::core::Interface;

use crate::core::UIElement;
use crate::errors::Error;
use crate::errors::Result;

pub trait IUIPattern : Sized {
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

impl IUIPattern for UIInvokePattern {
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

impl IUIPattern for UIAnnotationPattern {
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

impl IUIPattern for UICustomNavigationPattern {
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

impl IUIPattern for UIDockPattern {
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

impl IUIPattern for UIDragPattern {
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

impl IUIPattern for UIDropTargetPattern {
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

impl IUIPattern for UIExpandCollapsePattern {
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

impl IUIPattern for UIGridPattern {
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

impl IUIPattern for UIGridItemPattern {
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

impl IUIPattern for UIMultipleViewPattern {
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