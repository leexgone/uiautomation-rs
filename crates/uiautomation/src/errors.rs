use std::fmt::Display;

use windows::Win32::Foundation::GetLastError;
use windows::core::HRESULT;

/// Error caused by unknown reason.
pub const ERR_NONE: i32 = 0;
/// Error occurs when an element or object is not found.
pub const ERR_NOTFOUND: i32 = 1;
/// Error occurs when the operation is timeout.
pub const ERR_TIMEOUT: i32 = 2;
/// Error occurs when the element is inactive.
pub const ERR_INACTIVE: i32 = 3;
/// Error caused by unsupported or mismatched type
pub const ERR_TYPE: i32 = 4;
/// Error when a pointer is null.
pub const ERR_NULL_PTR:  i32 = 5;
/// Error format.
pub const ERR_FORMAT: i32 = 6;
/// Error invalid object.
pub const ERR_INVALID_OBJECT: i32 = 7;
/// Error already running.
pub const ERR_ALREADY_RUNNING: i32 = 8;

#[derive(Debug, PartialEq, Eq)]
pub struct Error {
    code: i32,
    message: String
}

impl Error {
    pub fn new(code: i32, message: &str) -> Error {
        Error {
            code,
            message: String::from(message)
        }
    }

    pub fn last_os_error() -> Error {
        let error = unsafe { GetLastError() };
        // let code: i32 = if (error.0 as i32) < 0 {
        //     error.0 as _
        // } else { 
        //     ((error.0 & 0x0000FFFF) | 0x80070000) as _
        // };

        // HRESULT(code).into()
        if let Err(e) = error {
            e.into()
        } else {
            HRESULT(0).into()
        }
    }

    pub fn code(&self) -> i32 {
        self.code
    }

    pub fn result(&self) -> Option<HRESULT> {
        if self.code < 0 {
            Some(HRESULT(self.code))
        } else {
            None
        }
    }

    pub fn message(&self) -> &str {
        self.message.as_str()
    }
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
        Self {
            code: e.code().0,
            message: e.message().to_string()
        }
    }
}

impl From<HRESULT> for Error {
    fn from(result: HRESULT) -> Self {
        Self {
            code: result.0,
            message: result.message().to_string()
        }
    }
}

impl From<String> for Error {
    fn from(message: String) -> Self {
        Error {
            code: 0,
            message
        }
    }
}

impl From<&str> for Error {
    fn from(message: &str) -> Self {
        Error {
            code: 0,
            message: String::from(message)
        }
    }
}

pub type Result<T> = core::result::Result<T, Error>;
