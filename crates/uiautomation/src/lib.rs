pub mod errors;
pub mod variants;
pub mod core;
pub mod patterns;
pub mod conditions;
pub mod controls;
pub mod actions;
pub mod inputs;

pub use crate::errors::Error;
pub use crate::errors::Result;

pub use crate::core::UIAutomation;
pub use crate::core::UIElement;
pub use crate::core::UITreeWalker;
pub use crate::core::UIMatcher;