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

pub trait ScrollItem {
    /// Scroll this item into view.
    fn scroll_into_view(&self) -> Result<()>;    
}