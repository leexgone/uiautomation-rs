pub mod errors;
pub mod types;
pub mod variants;
pub mod core;
pub mod patterns;
pub mod filters;
pub mod controls;
pub mod actions;
pub mod inputs;
pub mod processes;
pub mod graphics;

pub use self::errors::Error;
pub use self::errors::Result;

pub use self::core::UIAutomation;
pub use self::core::UIElement;
pub use self::core::UITreeWalker;
pub use self::core::UIMatcher;