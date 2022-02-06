use std::fmt::Display;

#[derive(Debug, PartialEq)]
pub struct Error {
    message: String
}

impl Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.message)
    }
}

impl std::error::Error for Error {}