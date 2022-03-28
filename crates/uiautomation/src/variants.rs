use std::fmt::Display;

use windows::Win32::Foundation::DECIMAL;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Com::SAFEARRAY;
use windows::Win32::System::Com::VARIANT;
use windows::Win32::System::Com::VARIANT_0_0_0;
use windows::Win32::System::Ole::VARENUM;
use windows::Win32::System::Ole::VT_CY;
use windows::Win32::System::Ole::VT_DATE;
use windows::Win32::System::Ole::VT_EMPTY;
use windows::Win32::System::Ole::VT_I1;
use windows::Win32::System::Ole::VT_I2;
use windows::Win32::System::Ole::VT_I4;
use windows::Win32::System::Ole::VT_I8;
use windows::Win32::System::Ole::VT_INT;
use windows::Win32::System::Ole::VT_NULL;
use windows::Win32::System::Ole::VT_R4;
use windows::Win32::System::Ole::VT_R8;
use windows::Win32::System::Ole::VT_UI1;
use windows::Win32::System::Ole::VT_UI2;
use windows::Win32::System::Ole::VT_UI4;
use windows::Win32::System::Ole::VT_UI8;
use windows::Win32::System::Ole::VT_UINT;
use windows::Win32::System::Ole::VT_VOID;
use windows::core::HRESULT;
use windows::core::IUnknown;

use crate::Error;
use crate::Result;
use crate::errors::ERR_TYPE;

/// enum type value for `Variant`
#[derive(Clone, PartialEq)]
pub enum Value {
    EMPTY,
    NULL,
    VOID,
    I1(i8),
    I2(i16),
    I4(i32),
    I8(i64),
    INT(i32),
    UI1(u8),
    UI2(u16),
    UI4(u32),
    UI8(u64),
    UINT(u32),
    R4(f32),
    R8(f64),
    CURRENCY(i64),
    DATE(f64),
    STRING(String),
    DISPATCH(IDispatch),
    ERROR(HRESULT),
    HRESULT(HRESULT),
    BOOL(bool),
    VARIANT(Variant),
    UNKNOWN(IUnknown),
    DECIMAL(DECIMAL),
    SAFEARRAY(SafeArray),
    ARRAY(SafeArray)
}

impl Display for Value {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Value::EMPTY => write!(f, "EMPTY"),
            Value::NULL => write!(f, "NULL"),
            Value::VOID => write!(f, "VOID"),
            Value::I1(value) => write!(f, "I1({})", value),
            Value::I2(value) => write!(f, "I2({})", value),
            Value::I4(value) => write!(f, "I4({})", value),
            Value::I8(value) => write!(f, "I8({})", value),
            Value::INT(value) => write!(f, "INT({})", value),
            Value::UI1(value) => write!(f, "UI1({})", value),
            Value::UI2(value) => write!(f, "UI2({})", value),
            Value::UI4(value) => write!(f, "UI4({})", value),
            Value::UI8(value) => write!(f, "UI8({})", value),
            Value::UINT(value) => write!(f, "UNIT({})", value),
            Value::R4(value) => write!(f, "R4({})", value),
            Value::R8(value) => write!(f, "R8({})", value),
            Value::CURRENCY(value) => write!(f, "CY({})", value),
            Value::DATE(value) => write!(f, "DATE({})", value),
            Value::STRING(value) => write!(f, "STRING({})", value),
            Value::DISPATCH(_) => write!(f, "DISPATCH"),
            Value::ERROR(value) => write!(f, "ERROR({})", value.0),
            Value::HRESULT(value) => write!(f, "HRESULT({})", value.0),
            Value::BOOL(value) => write!(f, "BOOL({})", value),
            Value::VARIANT(value) => write!(f, "VARIANT({})", value),
            Value::UNKNOWN(_) => write!(f, "UNKNOWN"),
            Value::DECIMAL(_) => write!(f, "DECIMAL"),
            Value::SAFEARRAY(value) => write!(f, "SAFEARRAY({})", value),
            Value::ARRAY(value) => write!(f, "ARRAY({})", value),
        }
    }
}

impl TryFrom<&Variant> for Value {
    type Error = Error;

    fn try_from(value: &Variant) -> Result<Self> {
        let vt = value.get_type();

        if vt == VT_EMPTY {
            Ok(Self::EMPTY)
        } else if vt == VT_NULL {
            Ok(Self::NULL)
        } else if vt == VT_VOID {
            Ok(Self::VOID)
        } else if vt == VT_I1 {
            let val = unsafe {
                value.get_data().bVal as i8
            }; 
            Ok(Self::I1(val))
        } else if vt == VT_I2 {
            let val = unsafe {
                value.get_data().iVal
            };
            Ok(Self::I2(val))
        } else if vt == VT_I4 {
            let val = unsafe {
                value.get_data().lVal
            };
            Ok(Self::I4(val))
        } else if vt == VT_I8 {
            let val = unsafe {
                value.get_data().llVal
            };
            Ok(Self::I8(val))
        } else if vt == VT_INT {
            let val = unsafe {
                value.get_data().lVal
            };
            Ok(Self::INT(val))
        } else if vt == VT_UI1 {
            let val = unsafe {
                value.get_data().bVal
            };
            Ok(Self::UI1(val))
        } else if vt == VT_UI2 {
            let val = unsafe {
                value.get_data().uiVal
            };
            Ok(Self::UI2(val))
        } else if vt == VT_UI4 {
            let val = unsafe {
                value.get_data().ulVal
            };
            Ok(Self::UI4(val))
        } else if vt == VT_UI8 {
            let val = unsafe {
                value.get_data().ullVal
            };
            Ok(Self::UI8(val))
        } else if vt == VT_UINT {
            let val = unsafe {
                value.get_data().uintVal
            };
            Ok(Self::UINT(val))
        } else if vt == VT_R4 {
            let val = unsafe {
                value.get_data().fltVal
            };
            Ok(Self::R4(val))
        } else if vt == VT_R8 {
            let val = unsafe {
                value.get_data().dblVal
            };
            Ok(Self::R8(val))
        } else if vt == VT_CY {
            let val = unsafe {
                value.get_data().cyVal.int64
            };
            Ok(Self::CURRENCY(val))
        } else if vt == VT_DATE {
            let val = unsafe {
                value.get_data().date
            };
            Ok(Self::DATE(val))
        }
        else {
            Err(Error::new(ERR_TYPE, ""))
        }
    }
}

/// A Wrapper for windows `VARIANT`
#[derive(Clone, PartialEq, Eq)]
pub struct Variant {
    value: VARIANT
}

impl Default for Variant {
    fn default() -> Self {
        Self { 
            value: Default::default() 
        }
    }
}

impl Variant {
    pub fn get_type(&self) -> VARENUM {
        let t = unsafe {
            self.value.Anonymous.Anonymous.vt as i32 
        }; 
        VARENUM(t)
    }

    pub(crate) unsafe fn get_data(&self) -> &VARIANT_0_0_0 {
        &self.value.Anonymous.Anonymous.Anonymous
    }
}

// impl TryInto<Value> for &Variant {
//     type Error = Error;

//     fn try_into(self) -> Result<Value> {
//         todo!()
//     }
// }

impl Display for Variant {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        if let Ok(val) = Value::try_from(self) {
            write!(f, "{}", val)
        } else {
            Err(std::fmt::Error {})
        }
    }
}

/// A Wrapper for windows `SAFEARRAY`
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SafeArray {
    array: SAFEARRAY
}

impl Display for SafeArray {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        todo!()
    }
}

#[cfg(test)]
mod tests {
    use std::mem::ManuallyDrop;

    use windows::Win32::Foundation::BSTR;
    use windows::Win32::System::Com::VARIANT;
    use windows::Win32::System::Com::VARIANT_0_0;
    use windows::Win32::System::Com::VARIANT_0_0_0;
    use windows::Win32::System::Ole::VT_BSTR;

    #[test]
    fn build_variant() {
        let s: BSTR = "hello,world!".into();
        let v = VARIANT {
            Anonymous: windows::Win32::System::Com::VARIANT_0 { 
                Anonymous: ManuallyDrop::new(VARIANT_0_0 { 
                    vt: VT_BSTR.0 as u16, 
                    wReserved1: 0, 
                    wReserved2: 0, 
                    wReserved3: 0, 
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(s)
                    }
                })
            }
        };
    }
}
