use windows::Win32::UI::Accessibility::WindowInteractionState;
use windows::Win32::UI::Accessibility::ZoomUnit;

use crate::Result;
use crate::UIElement;
use crate::variants::Variant;

/// Define a invokable action for ui element.
pub trait Invoke {
    /// Perform a click event on this control.
    fn invoke(&self) -> Result<()>;
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