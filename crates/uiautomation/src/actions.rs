use windows::Win32::UI::Accessibility::*;

use super::Result;
use super::UIElement;
use super::variants::Variant;

/// Define a invokable action for ui element.
pub trait Invoke {
    /// Perform a click event on this control.
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
    /// Select current item.
    fn select(&self) -> Result<()>;

    /// Add current item to selection.
    fn add_to_selection(&self) -> Result<()>;

    /// Remove current item from selection.
    fn remove_from_selection(&self) -> Result<()>;

    /// Determines whether this item is selected.
    fn is_selected(&self) -> Result<bool>;
}

/// Define a multiple view action for ui element.
pub trait MultipleView {
    /// Get supported view ids.
    fn get_supported_views(&self) -> Result<Vec<i32>>;

    /// Return the view name.
    /// 
    /// The `view` parameter is the id of the view. You can get all view ids by `get_supported_views()` function.
    fn get_view_name(&self, view: i32) -> Result<String>;

    /// Get the current view id.
    fn get_current_view(&self) -> Result<i32>;

    /// Set the current view by id.
    fn set_current_view(&self, view: i32) -> Result<()>;
}

/// Define a item container action for ui element.
pub trait ItemContainer {
    /// Find contained item by property value.
    /// 
    /// Search `UIElement` from the `start_ater` item, return the first item which property is specified value.
    fn find_item_by_property(&self, start_after: UIElement, property_id: i32, value: Variant) -> Result<UIElement>;    
}

/// Define a scroll action for ui element.
pub trait Scroll {
    /// Scroll the control.
    fn scroll(&self, horizontal_amount: ScrollAmount, vertical_amount: ScrollAmount) -> Result<()>;

    /// Set the scroll percent.
    fn set_scroll_percent(&self, horizontal_percent: f64, vertical_percent: f64) -> Result<()>;

    /// Get the horizontal scroll percent.
    fn get_horizontal_scroll_percent(&self) -> Result<f64>;

    /// Get the vertical scroll percent.
    fn get_vertical_scroll_percent(&self) -> Result<f64>;

    /// Get the horizontal view size.
    fn get_horizontal_view_size(&self) -> Result<f64>;

    /// Get the vertical view size.
    fn get_vertical_view_size(&self) -> Result<f64>;

    /// Get whether the control is horizontally scrollable.
    fn is_horizontally_scrollable(&self) -> Result<bool>;

    /// Get whether the control is vertically scrollable.
    fn is_vertically_scrollable(&self) -> Result<bool>;
}

/// Define a scroll item action for ui element.
pub trait ScrollItem {
    /// Scroll this item into view.
    fn scroll_into_view(&self) -> Result<()>;    
}

/// Define a window action for ui element.
pub trait Window {
    /// Close the window.
    fn close(&self) -> Result<()>;

    /// Wait for the window input idle.
    fn wait_for_input_idle(&self, milliseconds: i32) -> Result<bool>;

    /// Check whether the window is normal state.
    fn is_normal(&self) -> Result<bool>;

    /// Set the window state as normal.
    fn normal(&self) -> Result<()>;

    /// Check whether the window is able to be maximized.
    fn can_maximize(&self) -> Result<bool>;

    /// Chcek whether the window is maximized state. 
    fn is_maximized(&self) -> Result<bool>;

    /// Set the window state as maximized. 
    fn maximize(&self) -> Result<()>;

    /// Check whether the window is able to be minimized.
    fn can_minimize(&self) -> Result<bool>;

    /// Check whether the window is minimized state.
    fn is_minimized(&self) -> Result<bool>;

    /// Set the window state as minimized.
    fn minimize(&self) -> Result<()>;

    /// Check whether the window is model mode.
    fn is_modal(&self) -> Result<bool>;

    /// Check whether the window is topmost.
    fn is_topmost(&self) -> Result<bool>;

    /// Get the interaction state of the window.
    fn get_window_interaction_state(&self) -> Result<WindowInteractionState>;
}

/// Define a transform action for ui element.
pub trait Transform {
    /// Check whether the control is moveable.
    fn can_move(&self) -> Result<bool>;

    /// Move the control to the `(x, y)` point.
    fn move_to(&self, x: f64, y: f64) -> Result<()>;

    /// Check whether the control is resizable.
    fn can_resize(&self) -> Result<bool>;

    /// Resize the control size.
    fn resize(&self, width: f64, height: f64) -> Result<()>;

    /// Check whether the control is rotatable.
    fn can_rotate(&self) -> Result<bool>;

    /// Rotate the control by the `degrees`.
    fn rotate(&self, degrees: f64) -> Result<()>;

    /// Check whether the control is zoomable.
    fn can_zoom(&self) -> Result<bool>;

    /// Get the zoom level of the control.
    fn get_zoom_level(&self) -> Result<f64>;

    /// Get the minimum zoom size.
    fn get_zoom_minimum(&self) -> Result<f64>;

    /// Get the maximum zoom size.
    fn get_zoom_maximum(&self) -> Result<f64>;

    /// Zoom the control by `zoom_value` size.
    fn zoom(&self, zoom_value: f64) -> Result<()>;

    /// Zoom the control by `zoom_unit` unit.
    fn zoom_by_unit(&self, zoom_unit: ZoomUnit) -> Result<()>;
}

/// Define a value action for ui element.
pub trait Value {
    /// Set the edit value by `&str` type.
    fn set_value(&self, value: &str) -> Result<()>;

    /// Get the edit value as `String`.
    fn get_value(&self) -> Result<String>;

    /// Check whether the edit is readonly.
    fn is_readonly(&self) -> Result<bool>;
}

/// Define expand and collapse action for ui element.
pub trait ExpandCollapse {
    /// Expand the current control.
    fn expand(&self) -> Result<()>;

    /// Collapse the current control.
    fn collapse(&self) -> Result<()>;

    /// Get the state of the control.
    fn get_state(&self) -> Result<ExpandCollapseState>;
}

/// Define a toggle action for ui element.
pub trait Toggle {
    /// Get the toggle state.
    fn get_toggle_state(&self) -> Result<ToggleState>;

    /// Toggle the control.
    fn toggle(&self) -> Result<()>;
}

/// Define a grid action for ui element.
pub trait Grid {
    /// Get column count of the grid.
    fn get_column_count(&self) -> Result<i32>;

    /// Get row count of the grid.
    fn get_row_count(&self) -> Result<i32>;

    /// Get the item at [`row`, `column`] of the grid.
    fn get_item(&self, row: i32, column: i32) -> Result<UIElement>;
}

/// Define a table action for ui element.
pub trait Table {
    /// Get the row headers of the table.
    fn get_row_headers(&self) -> Result<Vec<UIElement>>;

    /// Get the column headers of the table.
    fn get_column_headers(&self) -> Result<Vec<UIElement>>;

    /// Get whether the row or column is major.
    fn get_row_or_column_major(&self) -> Result<RowOrColumnMajor>;
}