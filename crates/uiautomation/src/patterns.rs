use uiautomation_derive::pattern_as;
use uiautomation_derive::EnumConvert;
use uiautomation_derive::map_as;
use windows::core::Interface;
use windows::Win32::UI::Accessibility::IUIAutomationAnnotationPattern;
use windows::Win32::UI::Accessibility::IUIAutomationCustomNavigationPattern;
use windows::Win32::UI::Accessibility::IUIAutomationDockPattern;
use windows::Win32::UI::Accessibility::IUIAutomationDragPattern;
use windows::Win32::UI::Accessibility::IUIAutomationDropTargetPattern;
use windows::Win32::UI::Accessibility::IUIAutomationExpandCollapsePattern;
use windows::Win32::UI::Accessibility::IUIAutomationGridItemPattern;
use windows::Win32::UI::Accessibility::IUIAutomationGridPattern;
use windows::Win32::UI::Accessibility::IUIAutomationInvokePattern;
use windows::Win32::UI::Accessibility::IUIAutomationItemContainerPattern;
use windows::Win32::UI::Accessibility::IUIAutomationLegacyIAccessiblePattern;
use windows::Win32::UI::Accessibility::IUIAutomationMultipleViewPattern;
use windows::Win32::UI::Accessibility::IUIAutomationRangeValuePattern;
use windows::Win32::UI::Accessibility::IUIAutomationScrollItemPattern;
use windows::Win32::UI::Accessibility::IUIAutomationScrollPattern;
use windows::Win32::UI::Accessibility::IUIAutomationSelectionItemPattern;
use windows::Win32::UI::Accessibility::IUIAutomationSelectionPattern;
use windows::Win32::UI::Accessibility::IUIAutomationSelectionPattern2;
use windows::Win32::UI::Accessibility::IUIAutomationSpreadsheetItemPattern;
use windows::Win32::UI::Accessibility::IUIAutomationSpreadsheetPattern;
use windows::Win32::UI::Accessibility::IUIAutomationStylesPattern;
use windows::Win32::UI::Accessibility::IUIAutomationSynchronizedInputPattern;
use windows::Win32::UI::Accessibility::IUIAutomationTableItemPattern;
use windows::Win32::UI::Accessibility::IUIAutomationTablePattern;
use windows::Win32::UI::Accessibility::IUIAutomationTextChildPattern;
use windows::Win32::UI::Accessibility::IUIAutomationTextEditPattern;
use windows::Win32::UI::Accessibility::IUIAutomationTextPattern;
use windows::Win32::UI::Accessibility::IUIAutomationTextPattern2;
use windows::Win32::UI::Accessibility::IUIAutomationTextRange;
use windows::Win32::UI::Accessibility::IUIAutomationTextRange2;
use windows::Win32::UI::Accessibility::IUIAutomationTextRangeArray;
use windows::Win32::UI::Accessibility::IUIAutomationTogglePattern;
use windows::Win32::UI::Accessibility::IUIAutomationTransformPattern;
use windows::Win32::UI::Accessibility::IUIAutomationTransformPattern2;
use windows::Win32::UI::Accessibility::IUIAutomationValuePattern;
use windows::Win32::UI::Accessibility::IUIAutomationVirtualizedItemPattern;
use windows::Win32::UI::Accessibility::IUIAutomationWindowPattern;
use windows::Win32::UI::Accessibility::SynchronizedInputType;
use windows::core::BSTR;
use windows::core::IUnknown;
use windows_core::BOOL;

use crate::errors::Error;
use crate::Result;
use crate::UIElement;
use crate::types::AnnotationType;
use crate::types::DockPosition;
use crate::types::ExpandCollapseState;
use crate::types::NavigateDirection;
use crate::types::Point;
use crate::types::RowOrColumnMajor;
use crate::types::ScrollAmount;
use crate::types::StyleType;
use crate::types::SupportedTextSelection;
use crate::types::TextAttribute;
use crate::types::TextPatternRangeEndpoint;
use crate::types::TextUnit;
use crate::types::ToggleState;
use crate::types::UIProperty;
use crate::types::WindowInteractionState;
use crate::types::WindowVisualState;
use crate::types::ZoomUnit;
use crate::variants::SafeArray;
use crate::variants::Variant;

/// `UIPatternType` is an enum wrapper for `windows::Win32::UI::Accessibility::UIA_PATTERN_ID`.
/// 
/// Describes the named constants that identify Microsoft UI Automation control patterns.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::UIA_PATTERN_ID)]
pub enum UIPatternType {
    /// Identifies the Invoke control pattern.
    Invoke = 10000i32,
    /// Identifies the Selection control pattern.
    Selection = 10001i32,
    /// Identifies the Value control pattern.
    Value = 10002i32,
    /// Identifies the RangeValue control pattern.
    RangeValue = 10003i32,
    /// Identifies the Scroll control pattern.
    Scroll = 10004i32,
    /// Identifies the ExpandCollapse control pattern.
    ExpandCollapse = 10005i32,
    /// Identifies the Grid control pattern.
    Grid = 10006i32,
    /// Identifies the GridItem control pattern.
    GridItem = 10007i32,
    /// Identifies the MultipleView control pattern.
    MultipleView = 10008i32,
    /// Identifies the Window control pattern.
    Window = 10009i32,
    /// Identifies the SelectionItem control pattern.
    SelectionItem = 10010i32,
    /// Identifies the Dock control pattern.
    Dock = 10011i32,
    /// Identifies the Table control pattern.
    Table = 10012i32,
    /// Identifies the TableItem control pattern.
    TableItem = 10013i32,
    /// Identifies the Text control pattern.
    Text = 10014i32,
    /// Identifies the Toggle control pattern.
    Toggle = 10015i32,
    /// Identifies the Transform control pattern.
    Transform = 10016i32,
    /// Identifies the ScrollItem control pattern.
    ScrollItem = 10017i32,
    /// Identifies the LegacyIAccessible control pattern.
    LegacyIAccessible = 10018i32,
    /// Identifies the ItemContainer control pattern.
    ItemContainer = 10019i32,
    /// Identifies the VirtualizedItem control pattern.
    VirtualizedItem = 10020i32,
    /// Identifies the SynchronizedInput control pattern.
    SynchronizedInput = 10021i32,
    /// Identifies the ObjectModel control pattern. Supported starting with Windows 8.
    ObjectModel = 10022i32,
    /// Identifies the Annotation control pattern. Supported starting with Windows 8.
    Annotation = 10023i32,
    /// Identifies the second version of the Text control pattern. Supported starting with Windows 8.
    Text2 = 10024i32,
    /// Identifies the Styles control pattern. Supported starting with Windows 8.
    Styles = 10025i32,
    /// Identifies the Spreadsheet control pattern. Supported starting with Windows 8.
    Spreadsheet = 10026i32,
    /// Identifies the SpreadsheetItem control pattern. Supported starting with Windows 8.
    SpreadsheetItem = 10027i32,
    /// Identifies the second version of the Transform control pattern. Supported starting with Windows 8.
    Transform2 = 10028i32,
    /// Identifies the TextChild control pattern. Supported starting with Windows 8.
    TextChild = 10029i32,
    /// Identifies the Drag control pattern. Supported starting with Windows 8.
    Drag = 10030i32,
    /// Identifies the DropTarget control pattern. Supported starting with Windows 8.
    DropTarget = 10031i32,
    /// Identifies the TextEdit control pattern. Supported starting with Windows 8.1.
    TextEdit = 10032i32,
    /// Identifies the CustomNavigation control pattern. Supported starting with Windows 10.
    CustomNavigation = 10033i32    
}

/// `UIPattern` is the wrapper trait for patterns.
pub trait UIPattern {
    /// Defines the pattern type id.
    const TYPE: UIPatternType;
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Invoke)]
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

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Annotation)]
pub struct UIAnnotationPattern {
    pattern: IUIAutomationAnnotationPattern
}

impl UIAnnotationPattern {
    pub fn get_type(&self) -> Result<AnnotationType> {
        let id = unsafe {
            self.pattern.CurrentAnnotationTypeId()?
        };
        id.try_into()
    }

    pub fn get_cached_type(&self) -> Result<AnnotationType> {
        let id = unsafe {
            self.pattern.CachedAnnotationTypeId()?
        };
        id.try_into()
    }

    pub fn get_type_nane(&self) -> Result<String> {
        let name = unsafe {
            self.pattern.CurrentAnnotationTypeName()?
        };
        Ok(name.to_string())
    }

    pub fn get_cached_type_nane(&self) -> Result<String> {
        let name = unsafe {
            self.pattern.CachedAnnotationTypeName()?
        };
        Ok(name.to_string())
    }

    pub fn get_author(&self) -> Result<String> {
        let author = unsafe {
            self.pattern.CurrentAuthor()?
        };
        Ok(author.to_string())
    }

    pub fn get_cached_author(&self) -> Result<String> {
        let author = unsafe {
            self.pattern.CachedAuthor()?
        };
        Ok(author.to_string())
    }

    pub fn get_datetime(&self) -> Result<String> {
        let datetime = unsafe {
            self.pattern.CurrentDateTime()?
        };
        Ok(datetime.to_string())
    }

    pub fn get_cached_datetime(&self) -> Result<String> {
        let datetime = unsafe {
            self.pattern.CachedDateTime()?
        };
        Ok(datetime.to_string())
    }

    pub fn get_target(&self) -> Result<UIElement> {
        let target = unsafe {
            self.pattern.CurrentTarget()?
        };
        Ok(target.into())
    }

    pub fn get_cached_target(&self) -> Result<UIElement> {
        let target = unsafe {
            self.pattern.CachedTarget()?
        };
        Ok(target.into())
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::CustomNavigation)]
pub struct UICustomNavigationPattern {
    pattern: IUIAutomationCustomNavigationPattern
}

impl UICustomNavigationPattern {
    pub fn navigate(&self, direction: NavigateDirection) -> Result<UIElement> {
        let element = unsafe {
            self.pattern.Navigate(direction.into())?
        };
        Ok(element.into())
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Dock)]
pub struct UIDockPattern {
    pattern: IUIAutomationDockPattern
}

impl UIDockPattern {
    pub fn get_dock_position(&self) -> Result<DockPosition> {
        let pos = unsafe {
            self.pattern.CurrentDockPosition()?
        };
        // Ok(pos.into())
        pos.try_into()
    }

    pub fn get_cached_dock_position(&self) -> Result<DockPosition> {
        let pos = unsafe {
            self.pattern.CachedDockPosition()?
        };
        // Ok(pos.into())
        pos.try_into()
    }

    pub fn set_dock_position(&self, position: DockPosition) -> Result<()> {
        unsafe {
            self.pattern.SetDockPosition(position.into())?
        };
        Ok(())
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Drag)]
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

    pub fn is_cached_grabbed(&self) -> Result<bool> {
        let grabbed = unsafe {
            self.pattern.CachedIsGrabbed()?
        };
        Ok(grabbed.as_bool())
    }

    pub fn get_drop_effect(&self) -> Result<String> {
        let effect = unsafe {
            self.pattern.CurrentDropEffect()?
        };

        Ok(effect.to_string())
    }

    pub fn get_cached_drop_effect(&self) -> Result<String> {
        let effect = unsafe {
            self.pattern.CachedDropEffect()?
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

    pub fn get_cached_drop_effects(&self) -> Result<Vec<String>> {
        let effects = unsafe {
            self.pattern.CachedDropEffects()?
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

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::DropTarget)]
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

    pub fn get_cached_drop_target_effect(&self) -> Result<String> {
        let effect = unsafe {
            self.pattern.CachedDropTargetEffect()?
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

    pub fn get_cached_drop_target_effects(&self) -> Result<Vec<String>> {
        let effects = unsafe {
            self.pattern.CachedDropTargetEffects()?
        };

        let effects: SafeArray = effects.into();
        effects.try_into()
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::ExpandCollapse)]
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
        let state = unsafe {
            self.pattern.CurrentExpandCollapseState()?
        };
        // Ok(state.into())
        state.try_into()
    }

    pub fn get_cached_state(&self) -> Result<ExpandCollapseState> {
        let state = unsafe {
            self.pattern.CachedExpandCollapseState()?
        };
        // Ok(state.into())
        state.try_into()
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Grid)]
pub struct UIGridPattern {
    pattern: IUIAutomationGridPattern
}

impl UIGridPattern {
    pub fn get_column_count(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentColumnCount()?
        })
    }

    pub fn get_cached_column_count(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CachedColumnCount()?
        })
    }

    pub fn get_row_count(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentRowCount()?
        })
    }

    pub fn get_cached_row_count(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CachedRowCount()?
        })
    }

    pub fn get_item(&self, row: i32, column: i32) -> Result<UIElement> {
        let element = unsafe {
            self.pattern.GetItem(row, column)?
        };
        Ok(UIElement::from(element))
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::GridItem)]
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

    pub fn get_cached_containing_grid(&self) -> Result<UIElement> {
        let grid = unsafe {
            self.pattern.CachedContainingGrid()?
        };
        Ok(grid.into())
    }

    pub fn get_row(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentRow()?
        })
    }

    pub fn get_cached_row(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CachedRow()?
        })
    }

    pub fn get_column(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentColumn()?
        })
    }

    pub fn get_cached_column(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CachedColumn()?
        })
    }

    pub fn get_row_span(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentRowSpan()?
        })
    }

    pub fn get_cached_row_span(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CachedRowSpan()?
        })
    }

    pub fn get_column_span(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentColumnSpan()?
        })
    }

    pub fn get_cached_column_span(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CachedColumnSpan()?
        })
    }
}

/// A wrapper for `IUIAutomationItemContainerPattern`.
#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::ItemContainer)]
pub struct UIItemContainerPattern {
    pattern: IUIAutomationItemContainerPattern
}

impl UIItemContainerPattern {
    pub fn find_item_by_property(&self, start_after: UIElement, property: UIProperty, value: Variant) -> Result<UIElement> {
        // let val: VARIANT = value.into();
        let element = unsafe {
            self.pattern.FindItemByProperty(start_after.as_ref(), property.into(), value.as_ref())?
        };

        Ok(element.into())
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::LegacyIAccessible)]
pub struct UILegacyIAccessiblePattern {
    pattern: IUIAutomationLegacyIAccessiblePattern,
}

impl UILegacyIAccessiblePattern {
    pub fn select(&self, flags: i32) -> Result<()> {
        Ok(unsafe { self.pattern.Select(flags)? })
    }

    pub fn do_default_action(&self) -> Result<()> {
        Ok(unsafe { self.pattern.DoDefaultAction()? })
    }

    pub fn set_value(&self, value: &str) -> Result<()> {
        let value = BSTR::from(value);
        Ok(unsafe { self.pattern.SetValue(&value)? })
    }

    pub fn get_child_id(&self) -> Result<i32> {
        Ok(unsafe { self.pattern.CurrentChildId()? })
    }

    pub fn get_name(&self) -> Result<String> {
        let name = unsafe { self.pattern.CurrentName()? };
        Ok(name.to_string())
    }

    pub fn get_value(&self) -> Result<String> {
        let value = unsafe { self.pattern.CurrentValue()? };
        Ok(value.to_string())
    }

    pub fn get_description(&self) -> Result<String> {
        let desc = unsafe { self.pattern.CurrentDescription()? };
        Ok(desc.to_string())
    }

    pub fn get_role(&self) -> Result<u32> {
        Ok(unsafe { self.pattern.CurrentRole()? })
    }

    pub fn get_state(&self) -> Result<u32> {
        Ok(unsafe { self.pattern.CurrentState()? })
    }

    pub fn get_help(&self) -> Result<String> {
        let help = unsafe { self.pattern.CurrentHelp()? };
        Ok(help.to_string())
    }

    pub fn get_keyboard_shortcut(&self) -> Result<String> {
        let shortcut = unsafe { self.pattern.CurrentKeyboardShortcut()? };
        Ok(shortcut.to_string())
    }

    pub fn get_selection(&self) -> Result<Vec<UIElement>> {
        let selection = unsafe { self.pattern.GetCurrentSelection()? };
        crate::core::UIElement::to_elements(selection)
    }

    pub fn get_default_action(&self) -> Result<String> {
        let action = unsafe { self.pattern.CurrentDefaultAction()? };
        Ok(action.to_string())
    }

    pub fn get_cached_child_id(&self) -> Result<i32> {
        Ok(unsafe { self.pattern.CachedChildId()? })
    }

    pub fn get_cached_name(&self) -> Result<String> {
        let name = unsafe { self.pattern.CachedName()? };
        Ok(name.to_string())
    }

    pub fn get_cached_value(&self) -> Result<String> {
        let value = unsafe { self.pattern.CachedValue()? };
        Ok(value.to_string())
    }

    pub fn get_cached_description(&self) -> Result<String> {
        let desc = unsafe { self.pattern.CachedDescription()? };
        Ok(desc.to_string())
    }

    pub fn get_cached_role(&self) -> Result<u32> {
        Ok(unsafe { self.pattern.CachedRole()? })
    }

    pub fn get_cached_state(&self) -> Result<u32> {
        Ok(unsafe { self.pattern.CachedState()? })
    }

    pub fn get_cached_help(&self) -> Result<String> {
        let help = unsafe { self.pattern.CachedHelp()? };
        Ok(help.to_string())
    }

    pub fn get_cached_keyboard_shortcut(&self) -> Result<String> {
        let shortcut = unsafe { self.pattern.CachedKeyboardShortcut()? };
        Ok(shortcut.to_string())
    }

    pub fn get_cached_selection(&self) -> Result<Vec<UIElement>> {
        let selection = unsafe { self.pattern.GetCachedSelection()? };
        crate::core::UIElement::to_elements(selection)
    }

    pub fn get_cached_default_action(&self) -> Result<String> {
        let action = unsafe { self.pattern.CachedDefaultAction()? };
        Ok(action.to_string())
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::MultipleView)]
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

    pub fn get_cached_supported_views(&self) -> Result<Vec<i32>> {
        let views = unsafe {
            self.pattern.GetCachedSupportedViews()?
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

    pub fn get_cached_current_view(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CachedCurrentView()?
        })
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::RangeValue)]
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

    pub fn get_cached_value(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CachedValue()?
        })
    }

    pub fn is_readonly(&self) -> Result<bool> {
        let readonly = unsafe {
            self.pattern.CurrentIsReadOnly()?
        };
        Ok(readonly.as_bool())
    }

    pub fn is_cached_readonly(&self) -> Result<bool> {
        let readonly = unsafe {
            self.pattern.CachedIsReadOnly()?
        };
        Ok(readonly.as_bool())
    }

    pub fn get_maximum(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentMaximum()?
        })
    }

    pub fn get_cached_maximum(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CachedMaximum()?
        })
    }

    pub fn get_minimum(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentMinimum()?
        })
    }

    pub fn get_cached_minimum(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CachedMinimum()?
        })
    }

    pub fn get_large_change(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentLargeChange()?
        })
    }

    pub fn get_cached_large_change(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CachedLargeChange()?
        })
    }

    pub fn get_small_change(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentSmallChange()?
        })
    }

    pub fn get_cached_small_change(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CachedSmallChange()?
        })
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Scroll)]
pub struct UIScrollPattern {
    pattern: IUIAutomationScrollPattern
}

impl UIScrollPattern {
    pub fn scroll(&self, horizontal_amount: ScrollAmount, vertical_amount: ScrollAmount) -> Result<()> {
        Ok(unsafe {
            self.pattern.Scroll(horizontal_amount.into(), vertical_amount.into())?
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

    pub fn get_cached_horizontal_scroll_percent(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CachedHorizontalScrollPercent()?
        })
    }

    pub fn get_vertical_scroll_percent(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentVerticalScrollPercent()?
        })
    }

    pub fn get_cached_vertical_scroll_percent(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CachedVerticalScrollPercent()?
        })
    }

    pub fn get_horizontal_view_size(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentHorizontalViewSize()?
        })
    }

    pub fn get_cached_horizontal_view_size(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CachedHorizontalViewSize()?
        })
    }

    pub fn get_vertical_view_size(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CurrentVerticalViewSize()?
        })
    }

    pub fn get_cached_vertical_view_size(&self) -> Result<f64> {
        Ok(unsafe {
            self.pattern.CachedVerticalViewSize()?
        })
    }

    pub fn is_horizontally_scrollable(&self) -> Result<bool> {
        let scrollable = unsafe {
            self.pattern.CurrentHorizontallyScrollable()?
        };
        Ok(scrollable.as_bool())
    }

    pub fn is_cached_horizontally_scrollable(&self) -> Result<bool> {
        let scrollable = unsafe {
            self.pattern.CachedHorizontallyScrollable()?
        };
        Ok(scrollable.as_bool())
    }

    pub fn is_vertically_scrollable(&self) -> Result<bool> {
        let scrollable = unsafe {
            self.pattern.CurrentVerticallyScrollable()?
        };
        Ok(scrollable.as_bool())
    }

    pub fn is_cached_vertically_scrollable(&self) -> Result<bool> {
        let scrollable = unsafe {
            self.pattern.CachedVerticallyScrollable()?
        };
        Ok(scrollable.as_bool())
    }
}

// impl UIPattern for UIScrollPattern {
//     const TYPE: UIPatternType = UIPatternType::Scroll;
// }

// impl TryFrom<IUnknown> for UIScrollPattern {
//     type Error = Error;

//     fn try_from(value: IUnknown) -> Result<Self> {
//         let pattern: IUIAutomationScrollPattern = value.cast()?;
//         Ok(Self {
//             pattern
//         })
//     }
// }

// impl From<IUIAutomationScrollPattern> for UIScrollPattern {
//     fn from(pattern: IUIAutomationScrollPattern) -> Self {
//         Self {
//             pattern
//         }
//     }
// }

// impl Into<IUIAutomationScrollPattern> for UIScrollPattern {
//     fn into(self) -> IUIAutomationScrollPattern {
//         self.pattern
//     }
// }

// impl AsRef<IUIAutomationScrollPattern> for UIScrollPattern {
//     fn as_ref(&self) -> &IUIAutomationScrollPattern {
//         &self.pattern
//     }
// }

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::ScrollItem)]
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

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Selection)]
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

    pub fn get_cached_selection(&self) -> Result<Vec<UIElement>> {
        let elem_arr = unsafe {
            self.pattern.GetCachedSelection()?
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

    pub fn can_cached_select_multiple(&self) -> Result<bool> {
        let multiple = unsafe {
            self.pattern.CachedCanSelectMultiple()?
        };
        Ok(multiple.as_bool())
    }

    pub fn is_selection_required(&self) -> Result<bool> {
        let required = unsafe {
            self.pattern.CurrentIsSelectionRequired()?
        };
        Ok(required.as_bool())
    }

    pub fn is_cached_selection_required(&self) -> Result<bool> {
        let required = unsafe {
            self.pattern.CachedIsSelectionRequired()?
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

    pub fn get_cached_first_selected_item(&self) -> Result<UIElement> {
        let pattern2: IUIAutomationSelectionPattern2 = self.pattern.cast()?;
        let item = unsafe {
            pattern2.CachedFirstSelectedItem()?
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

    pub fn get_cached_last_selected_item(&self) -> Result<UIElement> {
        let pattern2: IUIAutomationSelectionPattern2 = self.pattern.cast()?;
        let item = unsafe {
            pattern2.CachedLastSelectedItem()?
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

    pub fn get_cached_current_selected_item(&self) -> Result<UIElement> {
        let pattern2: IUIAutomationSelectionPattern2 = self.pattern.cast()?;
        let item = unsafe {
            pattern2.CachedCurrentSelectedItem()?
        };
        Ok(item.into())
    }

    pub fn get_item_count(&self) -> Result<i32> {
        let pattern2: IUIAutomationSelectionPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CurrentItemCount()?
        })
    }

    pub fn get_cached_item_count(&self) -> Result<i32> {
        let pattern2: IUIAutomationSelectionPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CachedItemCount()?
        })
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::SelectionItem)]
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

    pub fn is_cached_selected(&self) -> Result<bool> {
        let selected = unsafe {
            self.pattern.CachedIsSelected()?
        };
        Ok(selected.as_bool())
    }

    pub fn get_selection_container(&self) -> Result<UIElement> {
        let container = unsafe {
            self.pattern.CurrentSelectionContainer()?
        };
        Ok(container.into())
    }

    pub fn get_cached_selection_container(&self) -> Result<UIElement> {
        let container = unsafe {
            self.pattern.CachedSelectionContainer()?
        };
        Ok(container.into())
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Spreadsheet)]
pub struct UISpreadsheetPattern {
    pattern: IUIAutomationSpreadsheetPattern
}

impl UISpreadsheetPattern {
    pub fn get_item_by_name(&self, name: &str) -> Result<UIElement> {
        let name = BSTR::from(name);
        let item = unsafe {
            self.pattern.GetItemByName(&name)?
        };
        Ok(item.into())
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::SpreadsheetItem)]
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

    pub fn get_cached_formula(&self) -> Result<String> {
        let formula = unsafe {
            self.pattern.CachedFormula()?
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

    pub fn get_cached_annotation_objects(&self) -> Result<Vec<UIElement>> {
        let elem_arr = unsafe {
            self.pattern.GetCachedAnnotationObjects()?
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

    pub fn get_cached_annotation_types(&self) -> Result<Vec<i32>> {
        let types = unsafe {
            self.pattern.GetCachedAnnotationTypes()?
        };

        let types: SafeArray = types.into();
        types.try_into()
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Styles)]
pub struct UIStylesPattern {
    pattern: IUIAutomationStylesPattern
}

impl UIStylesPattern {
    pub fn get_style(&self) -> Result<StyleType> {
        let style_id = unsafe {
            self.pattern.CurrentStyleId()?
        };
        style_id.try_into()
    }

    pub fn get_cached_style(&self) -> Result<StyleType> {
        let style_id = unsafe {
            self.pattern.CachedStyleId()?
        };
        style_id.try_into()
    }

    pub fn get_style_name(&self) -> Result<String> {
        let name = unsafe {
            self.pattern.CurrentStyleName()?
        };
        Ok(name.to_string())
    }

    pub fn get_cached_style_name(&self) -> Result<String> {
        let name = unsafe {
            self.pattern.CachedStyleName()?
        };
        Ok(name.to_string())
    }

    pub fn get_fill_color(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentFillColor()?
        })
    }

    pub fn get_cached_fill_color(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CachedFillColor()?
        })
    }

    pub fn get_fill_pattern_style(&self) -> Result<String> {
        let style = unsafe {
            self.pattern.CurrentFillPatternStyle()?
        };
        Ok(style.to_string())
    }

    pub fn get_cached_fill_pattern_style(&self) -> Result<String> {
        let style = unsafe {
            self.pattern.CachedFillPatternStyle()?
        };
        Ok(style.to_string())
    }

    pub fn get_fill_pattern_color(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CurrentFillPatternColor()?
        })
    }

    pub fn get_cached_fill_pattern_color(&self) -> Result<i32> {
        Ok(unsafe {
            self.pattern.CachedFillPatternColor()?
        })
    }

    pub fn get_shape(&self) -> Result<String> {
        let shape = unsafe {
            self.pattern.CurrentShape()?
        };
        Ok(shape.to_string())
    }

    pub fn get_cached_shape(&self) -> Result<String> {
        let shape = unsafe {
            self.pattern.CachedShape()?
        };
        Ok(shape.to_string())
    }

    pub fn get_extended_properties(&self) -> Result<String> {
        let properties = unsafe {
            self.pattern.CurrentExtendedProperties()?
        };
        Ok(properties.to_string())
    }

    pub fn get_cached_extended_properties(&self) -> Result<String> {
        let properties = unsafe {
            self.pattern.CachedExtendedProperties()?
        };
        Ok(properties.to_string())
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::SynchronizedInput)]
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

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Table)]
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

    pub fn get_cached_row_headers(&self) -> Result<Vec<UIElement>> {
        let headers = unsafe {
            self.pattern.GetCachedRowHeaders()?
        };

        UIElement::to_elements(headers)
    }

    pub fn get_column_headers(&self) -> Result<Vec<UIElement>> {
        let headers = unsafe {
            self.pattern.GetCurrentColumnHeaders()?
        };

        UIElement::to_elements(headers)
    }

    pub fn get_cached_column_headers(&self) -> Result<Vec<UIElement>> {
        let headers = unsafe {
            self.pattern.GetCachedColumnHeaders()?
        };

        UIElement::to_elements(headers)
    }

    pub fn get_row_or_column_major(&self) -> Result<RowOrColumnMajor> {
        let major = unsafe {
            self.pattern.CurrentRowOrColumnMajor()?
        };
        // Ok(major.into())
        major.try_into()
    }

    pub fn get_cached_row_or_column_major(&self) -> Result<RowOrColumnMajor> {
        let major = unsafe {
            self.pattern.CachedRowOrColumnMajor()?
        };
        // Ok(major.into())
        major.try_into()
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::TableItem)]
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

    pub fn get_cached_row_header_items(&self) -> Result<Vec<UIElement>> {
        let items = unsafe {
            self.pattern.GetCachedRowHeaderItems()?
        };

        UIElement::to_elements(items)
    }

    pub fn get_column_header_items(&self) -> Result<Vec<UIElement>> {
        let items = unsafe {
            self.pattern.GetCurrentColumnHeaderItems()?
        };

        UIElement::to_elements(items)
    }

    pub fn get_cached_column_header_items(&self) -> Result<Vec<UIElement>> {
        let items = unsafe {
            self.pattern.GetCachedColumnHeaderItems()?
        };

        UIElement::to_elements(items)
    }
}

#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::TextChild)]
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

/// A Wrapper for `IUIAutomationTextPattern` and `IUIAutomationTextPattern2`
#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Text)]
pub struct UITextPattern {
    pattern: IUIAutomationTextPattern
}

impl UITextPattern {
    pub fn get_range_from_point(&self, pt: Point) -> Result<UITextRange> {
        let range = unsafe {
            self.pattern.RangeFromPoint(pt.into())?
        };

        Ok(range.into())
    }

    pub fn get_range_from_child(&self, child: &UIElement) -> Result<UITextRange> {
        let range = unsafe {
            self.pattern.RangeFromChild(child.clone())?
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
        let selection = unsafe {
            self.pattern.SupportedTextSelection()?
        };
        selection.try_into()
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
        let range = unsafe {
            pattern2.GetCaretRange(&mut active)?
        };
        Ok((active.as_bool(), range.into()))
    }
}


/// A Wrapper for `IUIAutomationTextEditPattern`.
/// 
/// This type inherits from `UITextPattern`.
/// 
#[derive(Debug, Clone)]
// #[pattern_as(UIPatternType::TextEdit, windows::Win32::UI::Accessibility::IUIAutomationTextEditPattern)]
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
    const TYPE: UIPatternType = UIPatternType::TextEdit;
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
           self.range.CompareEndpoints(src_endpoint.into(), range.as_ref(), target_endpoint.into())?
       }) 
    }

    pub fn expand_to_enclosing_unit(&self, text_unit: TextUnit) -> Result<()> {
        Ok(unsafe {
            self.range.ExpandToEnclosingUnit(text_unit.into())?
        })
    }

    pub fn find_attribute(&self, attr: TextAttribute, value: Variant, backward: bool) -> Result<UITextRange> {
        let range = unsafe {
            self.range.FindAttribute(attr.into(), value.as_ref(), backward)?
        };
        Ok(range.into())
    }

    pub fn find_text(&self, text: &str, backward: bool, ignorecase: bool) -> Result<UITextRange> {
        let text: BSTR = text.into();
        let range = unsafe {
            self.range.FindText(&text, backward, ignorecase)?
        };
        Ok(range.into())
    }

    pub fn get_attribute_value(&self, attr: TextAttribute) -> Result<Variant> {
        let value = unsafe {
            self.range.GetAttributeValue(attr.into())?
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
            self.range.Move(unit.into(), count)?
        })
    }

    pub fn move_endpoint_by_unit(&self, endpoint: TextPatternRangeEndpoint, unit: TextUnit, count: i32) -> Result<i32> {
        Ok(unsafe {
            self.range.MoveEndpointByUnit(endpoint.into(), unit.into(), count)?
        })
    }

    pub fn move_endpoint_by_range(&self, src_endpoint: TextPatternRangeEndpoint, range: &UITextRange, target_endpoint: TextPatternRangeEndpoint) -> Result<()> {
        Ok(unsafe {
            self.range.MoveEndpointByRange(src_endpoint.into(), range.as_ref(), target_endpoint.into())?
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
#[pattern_as(UIPatternType::Toggle)]
pub struct UITogglePattern {
    pattern: IUIAutomationTogglePattern
}

impl UITogglePattern {
    pub fn get_toggle_state(&self) -> Result<ToggleState> {
        let state = unsafe {
            self.pattern.CurrentToggleState()?
        };
        // Ok(state.into())
        state.try_into()
    }

    pub fn get_cached_toggle_state(&self) -> Result<ToggleState> {
        let state = unsafe {
            self.pattern.CachedToggleState()?
        };
        // Ok(state.into())
        state.try_into()
    }

    pub fn toggle(&self) -> Result<()> {
        Ok(unsafe {
            self.pattern.Toggle()?
        })
    }
}

/// A wrapper for `IUIAutomationTransformPattern` and `IUIAutomationTransformPattern2`
#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Transform)]
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

    pub fn can_cached_move(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CachedCanMove()?
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

    pub fn can_cached_resize(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CachedCanResize()?
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

    pub fn can_cached_rotate(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CachedCanRotate()?
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

    pub fn can_cached_zoom(&self) -> Result<bool> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        let zoomable = unsafe {
            pattern2.CachedCanZoom()?
        };
        Ok(zoomable.as_bool())
    }

    pub fn get_zoom_level(&self) -> Result<f64> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CurrentZoomLevel()?
        })
    }

    pub fn get_cached_zoom_level(&self) -> Result<f64> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CachedZoomLevel()?
        })
    }

    pub fn get_zoom_minimum(&self) -> Result<f64> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CurrentZoomMinimum()?
        })
    }

    pub fn get_cached_zoom_minimum(&self) -> Result<f64> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CachedZoomMinimum()?
        })
    }

    pub fn get_zoom_maximum(&self) -> Result<f64> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CurrentZoomMaximum()?
        })
    }

    pub fn get_cached_zoom_maximum(&self) -> Result<f64> {
        let pattern2: IUIAutomationTransformPattern2 = self.pattern.cast()?;
        Ok(unsafe {
            pattern2.CachedZoomMaximum()?
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
            pattern2.ZoomByUnit(zoom_unit.into())?
        })
    }
}

/// A wrapper for `IUIAutomationValuePattern`.
#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Value)]
pub struct UIValuePattern {
    pattern: IUIAutomationValuePattern
}

impl UIValuePattern {
    pub fn set_value(&self, value: &str) -> Result<()> {
        let value = BSTR::from(value);
        Ok(unsafe {
            self.pattern.SetValue(&value)?
        })
    }

    pub fn get_value(&self) -> Result<String> {
        let value = unsafe {
            self.pattern.CurrentValue()?
        };
        Ok(value.to_string())
    }

    pub fn get_cached_value(&self) -> Result<String> {
        let value = unsafe {
            self.pattern.CachedValue()?
        };
        Ok(value.to_string())
   }

    pub fn is_readonly(&self) -> Result<bool> {
        let readonly = unsafe {
            self.pattern.CurrentIsReadOnly()?
        };
        Ok(readonly.as_bool())
    }

    pub fn cached_is_readonly(&self) -> Result<bool> {
        let readonly = unsafe {
            self.pattern.CachedIsReadOnly()?
        };
        Ok(readonly.as_bool())
    }

}

/// A wrapper for `IUIAutomationVirtualizedItemPattern`.
#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::VirtualizedItem)]
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

/// A wrapper for `IUIAutomationWindowPattern`.
#[derive(Debug, Clone)]
#[pattern_as(UIPatternType::Window)]
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
        let state = unsafe {
            self.pattern.CurrentWindowVisualState()?
        };
        state.try_into()
    }

    pub fn get_cached_window_visual_state(&self) -> Result<WindowVisualState> {
        let state = unsafe {
            self.pattern.CachedWindowVisualState()?
        };

        state.try_into()
    }

    pub fn set_window_visual_state(&self, state: WindowVisualState) -> Result<()> {
        Ok(unsafe {
            self.pattern.SetWindowVisualState(state.into())?
        })
    }

    pub fn can_maximize(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CurrentCanMaximize()?
        };
        Ok(ret.as_bool())
    }

    pub fn can_cached_maximize(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CachedCanMaximize()?
        };
        Ok(ret.as_bool())
    }

    pub fn can_minimize(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CurrentCanMinimize()?
        };
        Ok(ret.as_bool())
    }

    pub fn can_cached_minimize(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CachedCanMinimize()?
        };
        Ok(ret.as_bool())
    }

    pub fn is_modal(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CurrentIsModal()?
        };
        Ok(ret.as_bool())
    }

    pub fn cached_is_modal(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CachedIsModal()?
        };
        Ok(ret.as_bool())
    }

    pub fn is_topmost(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CurrentIsTopmost()?
        };
        Ok(ret.as_bool())
    }

    pub fn cached_is_topmost(&self) -> Result<bool> {
        let ret = unsafe {
            self.pattern.CachedIsTopmost()?
        };
        Ok(ret.as_bool())
    }

    pub fn get_window_interaction_state(&self) -> Result<WindowInteractionState> {
        let state = unsafe {
            self.pattern.CurrentWindowInteractionState()?
        };

        state.try_into()
    }

    pub fn get_cached_window_interaction_state(&self) -> Result<WindowInteractionState> {
        let state = unsafe {
            self.pattern.CachedWindowInteractionState()?
        };

        state.try_into()
    }
}

#[cfg(test)]
mod tests {
    use super::UIPatternType;

    #[test]
    fn test_uipatterntypes() {
        let t = UIPatternType::try_from(10000i32).unwrap();
        assert_eq!(t, UIPatternType::Invoke);
    }
}