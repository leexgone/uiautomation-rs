pub mod errors;
pub mod types;
pub mod variants;
pub mod core;
pub mod filters;

#[cfg(feature = "process")]
pub mod processes;
#[cfg(feature = "dialog")]
pub mod dialogs;
#[cfg(feature = "input")]
pub mod inputs;
#[cfg(feature = "pattern")]
pub mod patterns;
#[cfg(feature = "control")]
pub mod actions;
#[cfg(feature = "control")]
pub mod controls;
#[cfg(feature = "event")]
pub mod events;
#[cfg(feature = "clipboard")]
pub mod clipboards;

pub use self::errors::Error;
pub use self::errors::Result;

pub use self::core::UIAutomation;
pub use self::core::UIElement;
pub use self::core::UITreeWalker;
pub use self::core::UIMatcher;