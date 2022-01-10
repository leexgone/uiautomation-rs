use std::fmt::Error;

use windows::Win32::System::Com::VARIANT;

#[derive(Debug)]
pub enum Value<'a> {
    NULL,
    INT(i32),
    LONG(i64),
    STR(&'a str)
}
