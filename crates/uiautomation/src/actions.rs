use crate::Result;
use crate::UIElement;
use crate::patterns::UITextRange;
use crate::types::DockPosition;
use crate::types::ExpandCollapseState;
use crate::types::NavigateDirection;
use crate::types::Point;
use crate::types::RowOrColumnMajor;
use crate::types::ScrollAmount;
use crate::types::SupportedTextSelection;
use crate::types::ToggleState;
use crate::types::WindowInteractionState;
use crate::types::ZoomUnit;
use crate::variants::Variant;

/// Define a Invoke action for uielement.
pub trait Invoke {
    /// Invokes the action of a control, such as a button click.
    fn invoke(&self) -> Result<()>;
}

/// Define a selection action for ui element.
pub trait Selection {
    /// Retrieves the selected elements in the container.
    fn get_selection(&self) -> Result<Vec<UIElement>>;

    /// Indicates whether more than one item in the container can be selected at one time.
    fn can_select_multiple(&self) -> Result<bool>;

    /// Indicates whether at least one item must be selected at all times.
    fn is_selection_required(&self) -> Result<bool>;

    /// Gets an IUIAutomationElement object representing the first item in a group of selected items.
    fn get_first_selected_item(&self) -> Result<UIElement>;

    /// Gets an IUIAutomationElement object representing the last item in a group of selected items.
    fn get_last_selected_item(&self) -> Result<UIElement>;

    /// Gets an IUIAutomationElement object representing the currently selected item.
    fn get_current_selected_item(&self) -> Result<UIElement>;

    /// Gets an integer value indicating the number of selected items.
    fn get_item_count(&self) -> Result<i32>;
}

/// Define a selection item action for ui element.
pub trait SelectionItem {
    /// Clears any selected items and then selects the current element.
    fn select(&self) -> Result<()>;

    /// Adds the current element to the collection of selected items.
    fn add_to_selection(&self) -> Result<()>;

    /// Removes this element from the selection.
    fn remove_from_selection(&self) -> Result<()>;

    /// Indicates whether this item is selected.
    fn is_selected(&self) -> Result<bool>;

    /// Retrieves the element that supports IUIAutomationSelectionPattern and acts as the container for this item.
    fn get_selection_container(&self) -> Result<UIElement>;
}

/// Define a MultipleView action for uielement.
pub trait MultipleView {
    /// Retrieves a collection of control-specific view identifiers.
    fn get_supported_views(&self) -> Result<Vec<i32>>;

    /// Retrieves the name of a control-specific view.
    /// 
    /// The `view` parameter is the id of the view. You can get all view identifiers by `get_supported_views()` function.
    fn get_view_name(&self, view: i32) -> Result<String>;

    /// Retrieves the control-specific identifier of the current view of the control.
    fn get_current_view(&self) -> Result<i32>;

    /// Sets the view of the control.
    fn set_current_view(&self, view: i32) -> Result<()>;
}

/// Define a ItemContainer action for uielement.
pub trait ItemContainer {
    /// Retrieves an element within a containing element, based on a specified property value.
    /// 
    /// Search `UIElement` from the `start_ater` item, return the first item which property is specified value.
    fn find_item_by_property(&self, start_after: UIElement, property_id: i32, value: Variant) -> Result<UIElement>;    
}

/// Define a Scroll action for uielement.
pub trait Scroll {
    /// Scrolls the visible region of the content area horizontally and vertically.
    fn scroll(&self, horizontal_amount: ScrollAmount, vertical_amount: ScrollAmount) -> Result<()>;

    /// Sets the horizontal and vertical scroll positions as a percentage of the total content area within the UI Automation element.
    fn set_scroll_percent(&self, horizontal_percent: f64, vertical_percent: f64) -> Result<()>;

    /// Retrieves the horizontal scroll position.
    fn get_horizontal_scroll_percent(&self) -> Result<f64>;

    /// Retrieves the vertical scroll position.
    fn get_vertical_scroll_percent(&self) -> Result<f64>;

    /// Retrieves the horizontal size of the viewable region of a scrollable element.
    fn get_horizontal_view_size(&self) -> Result<f64>;

    /// Retrieves the vertical size of the viewable region of a scrollable element.
    fn get_vertical_view_size(&self) -> Result<f64>;

    /// Indicates whether the element can scroll horizontally.
    fn is_horizontally_scrollable(&self) -> Result<bool>;

    /// Indicates whether the element can scroll vertically.
    fn is_vertically_scrollable(&self) -> Result<bool>;
}

/// Define a ScrollItem action for uielement.
pub trait ScrollItem {
    /// Scrolls the content area of a container object to display the UI Automation element within the visible region (viewport) of the container.
    fn scroll_into_view(&self) -> Result<()>;    
}

/// Define a Window action for uielement.
pub trait Window {
    /// Close the window.
    fn close(&self) -> Result<()>;

    /// Causes the calling code to block for the specified time or until the associated process enters an idle state, whichever completes first.
    fn wait_for_input_idle(&self, milliseconds: i32) -> Result<bool>;

    /// Indicates whether the window is normal state.
    fn is_normal(&self) -> Result<bool>;

    /// Set the window state as normal.
    fn normal(&self) -> Result<()>;

    /// Indicates whether the window can be maximized.
    fn can_maximize(&self) -> Result<bool>;

    /// Chcek whether the window is maximized state. 
    fn is_maximized(&self) -> Result<bool>;

    /// Set the window state as maximized. 
    fn maximize(&self) -> Result<()>;

    /// Indicates whether the window can be minimized.
    fn can_minimize(&self) -> Result<bool>;

    /// Indicates whether the window is minimized state.
    fn is_minimized(&self) -> Result<bool>;

    /// Set the window state as minimized.
    fn minimize(&self) -> Result<()>;

    /// Indicates whether the window is modal.
    fn is_modal(&self) -> Result<bool>;

    /// Indicates whether the window is the topmost element in the z-order.
    fn is_topmost(&self) -> Result<bool>;

    /// Retrieves the current state of the window for the purposes of user interaction.
    fn get_window_interaction_state(&self) -> Result<WindowInteractionState>;
}

/// Define a Transform action for uielement.
pub trait Transform {
    /// Indicates whether the element can be moved.
    fn can_move(&self) -> Result<bool>;

    /// Moves the UI Automation element.
    fn move_to(&self, x: f64, y: f64) -> Result<()>;

    /// Indicates whether the element can be resized.
    fn can_resize(&self) -> Result<bool>;

    /// Resizes the UI Automation element.
    fn resize(&self, width: f64, height: f64) -> Result<()>;

    /// Indicates whether the element can be rotated.
    fn can_rotate(&self) -> Result<bool>;

    /// Rotates the UI Automation element.
    fn rotate(&self, degrees: f64) -> Result<()>;

    /// Indicates whether the control supports zooming of its viewport.
    fn can_zoom(&self) -> Result<bool>;

    /// Retrieves the zoom level of the control's viewport.
    fn get_zoom_level(&self) -> Result<f64>;

    /// Retrieves the minimum zoom level of the control's viewport.
    fn get_zoom_minimum(&self) -> Result<f64>;

    /// Retrieves the maximum zoom level of the control's viewport.
    fn get_zoom_maximum(&self) -> Result<f64>;

    /// Zooms the viewport of the control.
    fn zoom(&self, zoom_value: f64) -> Result<()>;

    /// Zooms the viewport of the control by the specified unit.
    fn zoom_by_unit(&self, zoom_unit: ZoomUnit) -> Result<()>;
}

/// Define a Value action for uielement.
pub trait Value {
    /// Sets the value of the element.
    fn set_value(&self, value: &str) -> Result<()>;

    /// Retrieves the value of the element.
    fn get_value(&self) -> Result<String>;

    /// Indicates whether the value of the element is read-only.
    fn is_readonly(&self) -> Result<bool>;
}

/// Define a ExpandCollapse action for uielement.
pub trait ExpandCollapse {
    /// Displays all child nodes, controls, or content of the element.
    fn expand(&self) -> Result<()>;

    /// Hides all child nodes, controls, or content of the element.
    fn collapse(&self) -> Result<()>;

    /// Retrieves a value that indicates the state, expanded or collapsed, of the element.
    fn get_state(&self) -> Result<ExpandCollapseState>;
}

/// Define a Toggle action for uielement.
pub trait Toggle {
    /// Retrieves the state of the control.
    fn get_toggle_state(&self) -> Result<ToggleState>;

    /// Cycles through the toggle states of the control.
    fn toggle(&self) -> Result<()>;
}

/// Define a Grid action for uielement.
pub trait Grid {
    /// The number of columns in the grid.
    fn get_column_count(&self) -> Result<i32>;

    /// Retrieves the number of rows in the grid.
    fn get_row_count(&self) -> Result<i32>;

    /// Retrieves a UI Automation element representing an item in the grid.
    fn get_item(&self, row: i32, column: i32) -> Result<UIElement>;
}

/// Define a Table action for uielement.
pub trait Table {
    /// Retrieves a collection of UI Automation elements representing all the row headers in a table.
    fn get_row_headers(&self) -> Result<Vec<UIElement>>;

    /// Retrieves a collection of UI Automation elements representing all the column headers in a table.
    fn get_column_headers(&self) -> Result<Vec<UIElement>>;

    /// Retrieves the primary direction of traversal for the table.
    fn get_row_or_column_major(&self) -> Result<RowOrColumnMajor>;
}

/// Define a CustomNavigation action for uielement.
pub trait CustomNavigation {
    /// Gets the next element in the specified direction within the logical UI tree.
    fn navigate(&self, direction: NavigateDirection) -> Result<UIElement>;
}

/// Define a GridItem action for uielement.
pub trait GridItem {
    /// Retrieves the element that contains the grid item.
    fn get_containing_grid(&self) -> Result<UIElement>;

    /// Retrieves the zero-based index of the row that contains the grid item.
    fn get_row(&self) -> Result<i32>;

    /// Retrieves the zero-based index of the column that contains the item.
    fn get_column(&self) -> Result<i32>;

    /// Retrieves the number of rows spanned by the grid item.
    fn get_row_span(&self) -> Result<i32>;

    /// Retrieves the number of columns spanned by the grid item.
    fn get_column_span(&self) -> Result<i32>;
}

/// Define a TableItem action for uielement.
pub trait TableItem {
    /// Retrieves the row headers associated with a table item or cell.
    fn get_row_header_items(&self) -> Result<Vec<UIElement>>;

    /// Retrieves the column headers associated with a table item or cell.
    fn get_column_header_items(&self) -> Result<Vec<UIElement>>;
}

/// Define a Text action for uielement.
pub trait Text {
    /// Retrieves the degenerate (empty) text range nearest to the specified screen coordinates.
    fn get_ragne_from_point(&self, pt: Point) -> Result<UITextRange>;

    /// Retrieves a text range enclosing a child element such as an image, hyperlink, Microsoft Excel spreadsheet, or other embedded object.
    fn get_range_from_child(&self, child: &UIElement) -> Result<UITextRange>;

    /// Retrieves a collection of text ranges that represents the currently selected text in a text-based control.
    fn get_selection(&self) -> Result<Vec<UITextRange>>;

    /// Retrieves an array of disjoint text ranges from a text-based control where each text range represents a contiguous span of visible text.
    fn get_visible_ranges(&self) -> Result<Vec<UITextRange>>;

    /// Retrieves a text range that encloses the main text of a document.
    fn get_document_range(&self) -> Result<UITextRange>;

    /// Retrieves a value that specifies the type of text selection that is supported by the control.
    fn get_supported_text_selection(&self) -> Result<SupportedTextSelection>;

    /// Retrieves a text range containing the text that is the target of the annotation associated with the specified annotation element.
    fn get_range_from_annotation(&self, annotation: &UIElement) -> Result<UITextRange>;

    /// Retrieves a zero-length text range at the location of the caret that belongs to the text-based control.
    fn get_caret_range(&self) -> Result<(bool, UITextRange)>;
}

/// Define a RangeValue action for uielement.
pub trait RangeValue {
    /// Sets the value of the control.
    fn set_value(&self, value: f64) -> Result<()>;

    /// Retrieves the value of the control.
    fn get_value(&self) -> Result<f64>;

    /// Indicates whether the value of the element can be changed.
    fn is_readonly(&self) -> Result<bool>;

    /// Retrieves the maximum value of the control.
    fn get_maximum(&self) -> Result<f64>;

    /// Retrieves the minimum value of the control.
    fn get_minimum(&self) -> Result<f64>;

    /// Retrieves the value that is added to or subtracted from the value of the control when a large change is made, such as when the PAGE DOWN key is pressed.
    fn get_large_change(&self) -> Result<f64>;

    /// Retrieves the value that is added to or subtracted from the value of the control when a small change is made, such as when an arrow key is pressed.
    fn get_small_change(&self) -> Result<f64>;
}

/// Define a Dock action for uielement.
pub trait Dock {
    /// Retrieves the dock position of this element within its docking container.
    fn get_dock_position(&self) -> Result<DockPosition>;

    /// Sets the dock position of this element.
    fn set_dock_position(&self, position: DockPosition) -> Result<()>;
}