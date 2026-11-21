pub mod core;
pub mod errors;
pub mod filters;
pub(crate) mod logs;
pub mod types;
pub mod variants;

#[cfg(feature = "control")]
pub mod actions;
#[cfg(feature = "clipboard")]
pub mod clipboards;
#[cfg(feature = "control")]
pub mod controls;
#[cfg(feature = "dialog")]
pub mod dialogs;
#[cfg(feature = "event")]
pub mod events;
#[cfg(feature = "input")]
pub mod inputs;
#[cfg(feature = "pattern")]
pub mod patterns;
#[cfg(feature = "process")]
pub mod processes;
#[cfg(feature = "screenshot")]
pub mod screenshots;

pub use self::errors::Error;
pub use self::errors::Result;

pub use self::core::UIAutomation;
pub use self::core::UIElement;
pub use self::core::UIMatcher;
pub use self::core::UITreeWalker;
