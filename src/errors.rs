use std::fmt::Display;

#[derive(Debug, PartialEq, Eq)]
pub struct Error {
    message: String
}

impl Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.message)
    }
}

impl std::error::Error for Error {
}

impl From<windows::core::Error> for Error {
    fn from(e: windows::core::Error) -> Self {
        e.to_string().into()
    }
}

impl From<String> for Error {
    fn from(message: String) -> Self {
        Error {
            message
        }
    }
}

impl From<&str> for Error {
    fn from(message: &str) -> Self {
        Error {
            message: String::from(message)
        }
    }
}
