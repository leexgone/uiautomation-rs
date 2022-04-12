use crate::Result;

/// Define a invokable control for ui element.
pub trait Invoke {
    /// Perform a click event on this control.
    fn invoke(&self) -> Result<()>;
}

/// Define a selection item control for ui element.
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