use std::fmt::Display;
use std::mem::ManuallyDrop;

use windows::Win32::Foundation::BSTR;
use windows::Win32::Foundation::DECIMAL;
use windows::Win32::System::Com::CY;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Com::SAFEARRAY;
use windows::Win32::System::Com::VARIANT;
use windows::Win32::System::Com::VARIANT_0;
use windows::Win32::System::Com::VARIANT_0_0;
use windows::Win32::System::Com::VARIANT_0_0_0;
use windows::Win32::System::Ole::SafeArrayCreate;
use windows::Win32::System::Ole::SafeArrayCreateVector;
use windows::Win32::System::Ole::VARENUM;
use windows::Win32::System::Ole::VT_ARRAY;
use windows::Win32::System::Ole::VT_BOOL;
use windows::Win32::System::Ole::VT_BSTR;
use windows::Win32::System::Ole::VT_CY;
use windows::Win32::System::Ole::VT_DATE;
use windows::Win32::System::Ole::VT_DECIMAL;
use windows::Win32::System::Ole::VT_DISPATCH;
use windows::Win32::System::Ole::VT_EMPTY;
use windows::Win32::System::Ole::VT_ERROR;
use windows::Win32::System::Ole::VT_HRESULT;
use windows::Win32::System::Ole::VT_I1;
use windows::Win32::System::Ole::VT_I2;
use windows::Win32::System::Ole::VT_I4;
use windows::Win32::System::Ole::VT_I8;
use windows::Win32::System::Ole::VT_INT;
use windows::Win32::System::Ole::VT_LPSTR;
use windows::Win32::System::Ole::VT_LPWSTR;
use windows::Win32::System::Ole::VT_NULL;
use windows::Win32::System::Ole::VT_R4;
use windows::Win32::System::Ole::VT_R8;
use windows::Win32::System::Ole::VT_SAFEARRAY;
use windows::Win32::System::Ole::VT_UI1;
use windows::Win32::System::Ole::VT_UI2;
use windows::Win32::System::Ole::VT_UI4;
use windows::Win32::System::Ole::VT_UI8;
use windows::Win32::System::Ole::VT_UINT;
use windows::Win32::System::Ole::VT_UNKNOWN;
use windows::Win32::System::Ole::VT_VARIANT;
use windows::Win32::System::Ole::VT_VOID;
use windows::Win32::System::Ole::VarBoolFromCy;
use windows::Win32::System::Ole::VarBoolFromDate;
use windows::Win32::System::Ole::VarBoolFromDec;
use windows::Win32::System::Ole::VarBoolFromDisp;
use windows::Win32::System::Ole::VarBoolFromI1;
use windows::Win32::System::Ole::VarBoolFromI2;
use windows::Win32::System::Ole::VarBoolFromI4;
use windows::Win32::System::Ole::VarBoolFromI8;
use windows::Win32::System::Ole::VarBoolFromR4;
use windows::Win32::System::Ole::VarBoolFromR8;
use windows::Win32::System::Ole::VarBoolFromStr;
use windows::Win32::System::Ole::VarBoolFromUI1;
use windows::Win32::System::Ole::VarBoolFromUI2;
use windows::Win32::System::Ole::VarBoolFromUI4;
use windows::Win32::System::Ole::VarBoolFromUI8;
use windows::Win32::System::Ole::VarBstrFromBool;
use windows::Win32::System::Ole::VarBstrFromCy;
use windows::Win32::System::Ole::VarBstrFromDate;
use windows::Win32::System::Ole::VarBstrFromDec;
use windows::Win32::System::Ole::VarBstrFromDisp;
use windows::Win32::System::Ole::VarBstrFromI1;
use windows::Win32::System::Ole::VarBstrFromI2;
use windows::Win32::System::Ole::VarBstrFromI4;
use windows::Win32::System::Ole::VarBstrFromI8;
use windows::Win32::System::Ole::VarBstrFromR4;
use windows::Win32::System::Ole::VarBstrFromR8;
use windows::Win32::System::Ole::VarBstrFromUI1;
use windows::Win32::System::Ole::VarBstrFromUI2;
use windows::Win32::System::Ole::VarBstrFromUI4;
use windows::Win32::System::Ole::VarBstrFromUI8;
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
    UNKNOWN(IUnknown),
    DISPATCH(IDispatch),
    ERROR(HRESULT),
    HRESULT(HRESULT),
    BOOL(bool),
    VARIANT(Variant),
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
            Value::UNKNOWN(_) => write!(f, "UNKNOWN"),
            Value::DISPATCH(_) => write!(f, "DISPATCH"),
            Value::ERROR(value) => write!(f, "ERROR({})", value.0),
            Value::HRESULT(value) => write!(f, "HRESULT({})", value.0),
            Value::BOOL(value) => write!(f, "BOOL({})", value),
            Value::VARIANT(value) => write!(f, "VARIANT({})", value),
            Value::DECIMAL(_) => write!(f, "DECIMAL"),
            Value::SAFEARRAY(value) => write!(f, "SAFEARRAY({})", value),
            Value::ARRAY(value) => write!(f, "ARRAY({})", value),
        }
    }
}

// impl TryFrom<&Variant> for Value {
//     type Error = Error;

//     fn try_from(value: &Variant) -> Result<Self> {
//         let vt = value.get_type();

//         if vt == VT_EMPTY {
//             Ok(Self::EMPTY)
//         } else if vt == VT_NULL {
//             Ok(Self::NULL)
//         } else if vt == VT_VOID {
//             Ok(Self::VOID)
//         } else if vt == VT_I1 {
//             let val = unsafe {
//                 value.get_data().bVal as i8
//             }; 
//             Ok(Self::I1(val))
//         } else if vt == VT_I2 {
//             let val = unsafe {
//                 value.get_data().iVal
//             };
//             Ok(Self::I2(val))
//         } else if vt == VT_I4 {
//             let val = unsafe {
//                 value.get_data().lVal
//             };
//             Ok(Self::I4(val))
//         } else if vt == VT_I8 {
//             let val = unsafe {
//                 value.get_data().llVal
//             };
//             Ok(Self::I8(val))
//         } else if vt == VT_INT {
//             let val = unsafe {
//                 value.get_data().lVal
//             };
//             Ok(Self::INT(val))
//         } else if vt == VT_UI1 {
//             let val = unsafe {
//                 value.get_data().bVal
//             };
//             Ok(Self::UI1(val))
//         } else if vt == VT_UI2 {
//             let val = unsafe {
//                 value.get_data().uiVal
//             };
//             Ok(Self::UI2(val))
//         } else if vt == VT_UI4 {
//             let val = unsafe {
//                 value.get_data().ulVal
//             };
//             Ok(Self::UI4(val))
//         } else if vt == VT_UI8 {
//             let val = unsafe {
//                 value.get_data().ullVal
//             };
//             Ok(Self::UI8(val))
//         } else if vt == VT_UINT {
//             let val = unsafe {
//                 value.get_data().uintVal
//             };
//             Ok(Self::UINT(val))
//         } else if vt == VT_R4 {
//             let val = unsafe {
//                 value.get_data().fltVal
//             };
//             Ok(Self::R4(val))
//         } else if vt == VT_R8 {
//             let val = unsafe {
//                 value.get_data().dblVal
//             };
//             Ok(Self::R8(val))
//         } else if vt == VT_CY {
//             let val = unsafe {
//                 value.get_data().cyVal.int64
//             };
//             Ok(Self::CURRENCY(val))
//         } else if vt == VT_DATE {
//             let val = unsafe {
//                 value.get_data().date
//             };
//             Ok(Self::DATE(val))
//         } else if vt == VT_BSTR || vt == VT_LPSTR {
//             let val = unsafe {
//                 value.get_data().bstrVal.to_string()
//             };
//             Ok(Self::STRING(val))
//         } else if vt == VT_LPSTR {
//             let val = unsafe {
//                 if value.get_data().pcVal.is_null() {
//                     String::from("")
//                 } else {
//                     let lpstr = value.get_data().pcVal.0;
//                     let mut end = lpstr;
//                     while *end != 0 {
//                         end = end.add(1);
//                     };
//                     String::from_utf8_lossy(std::slice::from_raw_parts(lpstr, end.offset_from(lpstr) as _)).into()
//                 }
//             };

//             Ok(Self::STRING(val))
//         } else if vt == VT_DISPATCH {
//             let val = unsafe {
//                 if let Some(ref disp) = *value.get_data().ppdispVal {
//                     Self::DISPATCH(disp.clone())
//                 } else {
//                     Self::NULL
//                 }
//             };
//             Ok(val)
//         } else if vt == VT_UNKNOWN {
//             let val = unsafe {
//                 if let Some(ref unkown) = *value.get_data().ppunkVal {
//                     Self::UNKNOWN(unkown.clone())
//                 } else {
//                     Self::NULL
//                 }
//             };
//             Ok(val)
//         } else if vt == VT_ERROR {
//             let val = unsafe {
//                 value.get_data().intVal
//             };
//             Ok(Self::HRESULT(HRESULT(val)))
//         } else if vt == VT_HRESULT {
//             let val = unsafe {
//                 value.get_data().intVal
//             };
//             Ok(Self::HRESULT(HRESULT(val)))
//         } else if vt == VT_BOOL {
//             let val = unsafe {
//                 value.get_data().__OBSOLETE__VARIANT_BOOL != 0
//             };
//             Ok(Self::BOOL(val))
//         } else if vt == VT_VARIANT {
//             let val = unsafe {
//                 (*value.get_data().pvarVal).clone()
//             };
//             Ok(Self::VARIANT(val.into()))
//         } else if vt == VT_DECIMAL {
//             let val = unsafe {
//                 (*value.get_data().pdecVal).clone()
//             };
//             Ok(Self::DECIMAL(val))
//         } else if vt == VT_SAFEARRAY || vt == VT_ARRAY {
//             let arr = unsafe {
//                 (*value.get_data().parray).clone()
//             };
//             Ok(Self::SAFEARRAY(arr.into()))
//         } else {
//             Err(Error::new(ERR_TYPE, ""))
//         }
//     }
// }

/// A Wrapper for windows `VARIANT`
#[derive(Clone, PartialEq, Eq, Default)]
pub struct Variant {
    value: VARIANT
}

impl Variant {
    /// Create a null variant.
    fn new_null(vt: VARENUM) -> Variant {
        let mut val = VARIANT_0_0::default();
        val.vt = vt.0 as u16;

        let variant = VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(val)
            }
        };

        variant.into()
    }

    /// Create a `Variant` from `vt` and `value`.
    fn new(vt: VARENUM, value: VARIANT_0_0_0) -> Variant {
        let variant = VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: vt.0 as u16,
                    wReserved1: 0,
                    wReserved2: 0,
                    wReserved3: 0,
                    Anonymous: value
                })
            }
        };

        variant.into()
    }

    /// Retrieve the variant type as `i32`.
    fn vt(&self) -> i32 {
        unsafe {
            self.value.Anonymous.Anonymous.vt as i32
        }
    }

    /// Retrieve the variant type as `VARENUM`.
    pub fn get_type(&self) -> VARENUM {
        VARENUM(self.vt())
    }

    /// Retrieve the data of the variant.
    pub(crate) unsafe fn get_data(&self) -> &VARIANT_0_0_0 {
        &self.value.Anonymous.Anonymous.Anonymous
    }

    /// Try to get value.
    pub fn get_value(&self) -> Result<Value> {
        self.try_into()
    }

    /// Check whether the variant is null.
    /// 
    /// Return `true` when vt is `VT_EMPTY`, `VT_NULL` or `VT_VOID`.
    pub fn is_null(&self) -> bool {
        let vt = self.vt();
        vt == VT_EMPTY.0 || vt == VT_NULL.0 || vt == VT_VOID.0
    }

    /// Check whether the variant is string.
    /// 
    /// Return `true` when vt is `VT_BSTR`, `VT_LPWSTR` or `VT_LPSTR`.
    pub fn is_string(&self) -> bool {
        let vt = self.vt();
        vt == VT_BSTR.0 || vt == VT_LPWSTR.0 || vt == VT_LPSTR.0
    }

    /// Try to get string value.
    /// 
    /// Return `String` value when vt is `VT_BSTR`, `VT_LPWSTR` or `VT_LPSTR`.
    pub fn get_string(&self) -> Result<String> {
        let value = self.get_value()?;
        match value {
            Value::STRING(str) => Ok(str),
            _ => Err(Error::new(ERR_TYPE, "Error Variant Type"))
        }
    }

    /// Check whether the variant is array.
    /// 
    /// Return `true` when vt is `VT_SAFEARRAY` or `VT_ARRAY`.
    pub fn is_array(&self) -> bool {
        let vt = self.vt();
        vt == VT_SAFEARRAY.0 || vt == VT_ARRAY.0
    }

    /// Try to get array value.
    /// 
    /// Return `SafeArray` value when vt is `VT_SAFEARRAY` or `VT_ARRAY`.
    pub fn get_array(&self) -> Result<SafeArray> {
        let value = self.get_value()?;
        match value {
            Value::SAFEARRAY(arr) => Ok(arr),
            Value::ARRAY(arr) => Ok(arr),
            _ => Err(Error::new(ERR_TYPE, "Error Variant Type"))
        }
    }
}

impl From<VARIANT> for Variant {
    fn from(value: VARIANT) -> Self {
        Self {
            value
        }
    }
}

impl Into<VARIANT> for Variant {
    fn into(self) -> VARIANT {
        self.value
    }
}

impl AsRef<VARIANT> for Variant {
    fn as_ref(&self) -> &VARIANT {
        &self.value
    }
}

impl Display for Variant {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        if let Ok(val) = self.get_value() {
            write!(f, "{}", val)
        } else {
            Err(std::fmt::Error {})
        }
    }
}

impl From<Value> for Variant {
    fn from(value: Value) -> Self {
        match value {
            Value::EMPTY => Variant::new_null(VT_EMPTY),
            Value::NULL => Variant::new_null(VT_NULL),
            Value::VOID => Variant::new_null(VT_VOID),
            Value::I1(v) => Variant::new(VT_I1, VARIANT_0_0_0 { bVal: v as u8 }),
            Value::I2(v) => Variant::new(VT_I2, VARIANT_0_0_0 { iVal: v }),
            Value::I4(v) => Variant::new(VT_I4, VARIANT_0_0_0 { lVal: v }),
            Value::I8(v) => Variant::new(VT_I8, VARIANT_0_0_0 { llVal: v }),
            Value::INT(v) => Variant::new(VT_INT, VARIANT_0_0_0 { lVal: v }),
            Value::UI1(v) => Variant::new(VT_UI1, VARIANT_0_0_0 { bVal: v }),
            Value::UI2(v) => Variant::new(VT_UI2, VARIANT_0_0_0 { uiVal: v }),
            Value::UI4(v) => Variant::new(VT_UI4, VARIANT_0_0_0 { ulVal: v }),
            Value::UI8(v) => Variant::new(VT_UI8, VARIANT_0_0_0 { ullVal: v }),
            Value::UINT(v) => Variant::new(VT_UINT, VARIANT_0_0_0 { uintVal: v }),
            Value::R4(v) => Variant::new(VT_R4, VARIANT_0_0_0 { fltVal: v }),
            Value::R8(v) => Variant::new(VT_R8, VARIANT_0_0_0 { dblVal: v }),
            Value::CURRENCY(v) => Variant::new(VT_CY, VARIANT_0_0_0 { cyVal: CY { int64: v} }),
            Value::DATE(v) => Variant::new(VT_DATE, VARIANT_0_0_0 { date: v }),
            Value::STRING(v) => Variant::new(VT_BSTR, VARIANT_0_0_0 { bstrVal: ManuallyDrop::new(BSTR::from(v)) }),
            Value::UNKNOWN(v) => Variant::new(VT_UNKNOWN, VARIANT_0_0_0 { punkVal: ManuallyDrop::new(Some(v)) }),
            Value::DISPATCH(v) => Variant::new(VT_DISPATCH, VARIANT_0_0_0 { pdispVal: ManuallyDrop::new(Some(v)) }),
            Value::ERROR(v) => Variant::new(VT_ERROR, VARIANT_0_0_0 { intVal: v.0 }),
            Value::HRESULT(v) => Variant::new(VT_HRESULT, VARIANT_0_0_0 { intVal: v.0 }),
            Value::BOOL(v) => Variant::new(VT_BOOL, VARIANT_0_0_0 { boolVal: if v { 0xffff } else { 0x000 } as i16 }),
            Value::VARIANT(mut v) => Variant::new(VT_VARIANT, VARIANT_0_0_0 { pvarVal: &mut v.value }),
            Value::DECIMAL(mut v) => Variant::new(VT_DECIMAL, VARIANT_0_0_0 { pdecVal: &mut v }),
            Value::SAFEARRAY(mut v) => Variant::new(VT_SAFEARRAY, VARIANT_0_0_0 { parray: &mut v.array }),
            Value::ARRAY(mut v) => Variant::new(VT_SAFEARRAY, VARIANT_0_0_0 { parray: &mut v.array }),
        }
    }
}

impl TryInto<Value> for &Variant {
    type Error = Error;

    fn try_into(self) -> Result<Value> {
        let vt = self.vt();

        if vt == VT_EMPTY.0 {
            Ok(Value::EMPTY)
        } else if vt == VT_NULL.0 {
            Ok(Value::NULL)
        } else if vt == VT_VOID.0 {
            Ok(Value::VOID)
        } else if vt == VT_I1.0 {
            let val = unsafe {
                self.get_data().bVal as i8
            }; 
            Ok(Value::I1(val))
        } else if vt == VT_I2.0 {
            let val = unsafe {
                self.get_data().iVal
            };
            Ok(Value::I2(val))
        } else if vt == VT_I4.0 {
            let val = unsafe {
                self.get_data().lVal
            };
            Ok(Value::I4(val))
        } else if vt == VT_I8.0 {
            let val = unsafe {
                self.get_data().llVal
            };
            Ok(Value::I8(val))
        } else if vt == VT_INT.0 {
            let val = unsafe {
                self.get_data().lVal
            };
            Ok(Value::INT(val))
        } else if vt == VT_UI1.0 {
            let val = unsafe {
                self.get_data().bVal
            };
            Ok(Value::UI1(val))
        } else if vt == VT_UI2.0 {
            let val = unsafe {
                self.get_data().uiVal
            };
            Ok(Value::UI2(val))
        } else if vt == VT_UI4.0 {
            let val = unsafe {
                self.get_data().ulVal
            };
            Ok(Value::UI4(val))
        } else if vt == VT_UI8.0 {
            let val = unsafe {
                self.get_data().ullVal
            };
            Ok(Value::UI8(val))
        } else if vt == VT_UINT.0 {
            let val = unsafe {
                self.get_data().uintVal
            };
            Ok(Value::UINT(val))
        } else if vt == VT_R4.0 {
            let val = unsafe {
                self.get_data().fltVal
            };
            Ok(Value::R4(val))
        } else if vt == VT_R8.0 {
            let val = unsafe {
                self.get_data().dblVal
            };
            Ok(Value::R8(val))
        } else if vt == VT_CY.0 {
            let val = unsafe {
                self.get_data().cyVal.int64
            };
            Ok(Value::CURRENCY(val))
        } else if vt == VT_DATE.0 {
            let val = unsafe {
                self.get_data().date
            };
            Ok(Value::DATE(val))
        } else if vt == VT_BSTR.0 || vt == VT_LPSTR.0 {
            let val = unsafe {
                self.get_data().bstrVal.to_string()
            };
            Ok(Value::STRING(val))
        } else if vt == VT_LPSTR.0 {
            let val = unsafe {
                if self.get_data().pcVal.is_null() {
                    String::from("")
                } else {
                    let lpstr = self.get_data().pcVal.0;
                    let mut end = lpstr;
                    while *end != 0 {
                        end = end.add(1);
                    };
                    String::from_utf8_lossy(std::slice::from_raw_parts(lpstr, end.offset_from(lpstr) as _)).into()
                }
            };

            Ok(Value::STRING(val))
        } else if vt == VT_DISPATCH.0 {
            let val = unsafe {
                if let Some(ref disp) = *self.get_data().ppdispVal {
                    Value::DISPATCH(disp.clone())
                } else {
                    Value::NULL
                }
            };
            Ok(val)
        } else if vt == VT_UNKNOWN.0 {
            let val = unsafe {
                if let Some(ref unkown) = *self.get_data().ppunkVal {
                    Value::UNKNOWN(unkown.clone())
                } else {
                    Value::NULL
                }
            };
            Ok(val)
        } else if vt == VT_ERROR.0 {
            let val = unsafe {
                self.get_data().intVal
            };
            Ok(Value::HRESULT(HRESULT(val)))
        } else if vt == VT_HRESULT.0 {
            let val = unsafe {
                self.get_data().intVal
            };
            Ok(Value::HRESULT(HRESULT(val)))
        } else if vt == VT_BOOL.0 {
            let val = unsafe {
                self.get_data().__OBSOLETE__VARIANT_BOOL != 0
            };
            Ok(Value::BOOL(val))
        } else if vt == VT_VARIANT.0 {
            let val = unsafe {
                (*self.get_data().pvarVal).clone()
            };
            Ok(Value::VARIANT(val.into()))
        } else if vt == VT_DECIMAL.0 {
            let val = unsafe {
                (*self.get_data().pdecVal).clone()
            };
            Ok(Value::DECIMAL(val))
        } else if vt == VT_SAFEARRAY.0 || vt == VT_ARRAY.0 {
            let arr = unsafe {
                (*self.get_data().parray).clone()
            };
            Ok(Value::SAFEARRAY(arr.into()))
        } else {
            Err(Error::new(ERR_TYPE, ""))
        }
    }
}

impl TryInto<Value> for Variant {
    type Error = Error;

    fn try_into(self) -> Result<Value> {
        (&self).try_into()
    }
}

impl From<bool> for Variant {
    fn from(value: bool) -> Self {
        Value::BOOL(value).into()
    }
}

impl TryInto<bool> for Variant {
    type Error = Error;

    fn try_into(self) -> Result<bool> {
        let vt = self.vt();
        let val: i16 = unsafe { if vt == VT_BOOL.0 {
                self.get_data().boolVal
            } else if vt == VT_CY.0 {
                VarBoolFromCy(self.get_data().cyVal)?
            } else if vt == VT_DATE.0 {
                VarBoolFromDate(self.get_data().date)?
            } else if vt == VT_DECIMAL.0 {
                VarBoolFromDec(self.get_data().pdecVal)?
            } else if vt == VT_I1.0 {
                VarBoolFromI1(self.get_data().cVal)?
            } else if vt == VT_I2.0 {
                VarBoolFromI2(self.get_data().iVal)?
            } else if vt == VT_I4.0 || vt == VT_INT.0 {
                VarBoolFromI4(self.get_data().lVal)?
            } else if vt == VT_I8.0 {
                VarBoolFromI8(self.get_data().llVal)?
            } else if vt == VT_R4.0 {
                VarBoolFromR4(self.get_data().fltVal)?
            } else if vt == VT_R8.0 {
                VarBoolFromR8(self.get_data().dblVal)?
            } else if vt == VT_BSTR.0 || vt == VT_LPWSTR.0 || vt == VT_LPSTR.0 {
                let str = self.get_string()?;
                VarBoolFromStr(str, 0, 0)?
            } else if vt == VT_UI1.0 {
                VarBoolFromUI1(self.get_data().bVal)?
            } else if vt == VT_UI2.0 {
                VarBoolFromUI2(self.get_data().uiVal)?
            } else if vt == VT_UI4.0 || vt == VT_UINT.0 {
                VarBoolFromUI4(self.get_data().ulVal)?
            } else if vt == VT_UI8.0 {
                VarBoolFromUI8(self.get_data().ullVal)?
            } else if vt == VT_DISPATCH.0 {
                if let Some(ref disp) = *self.get_data().pdispVal {
                    VarBoolFromDisp(disp, 0)?
                } else {
                    0i16
                }
            } else {
                return Err(Error::new(ERR_TYPE, "Error Variant Type"));
            }
        };
        Ok(val != 0)
    }
}

impl From<&str> for Variant {
    fn from(value: &str) -> Self {
        Value::STRING(value.into()).into()
    }
}

impl From<String> for Variant {
    fn from(value: String) -> Self {
        value.as_str().into()
    }
}

impl From<&String> for Variant {
    fn from(value: &String) -> Self {
        value.as_str().into()
    }
}

impl TryInto<String> for &Variant {
    type Error = Error;

    fn try_into(self) -> Result<String> {
        if self.is_string() {
            self.get_string()
        } else {
            let vt = self.vt();
            let str: BSTR = unsafe {
                if vt == VT_BOOL.0 {
                    VarBstrFromBool(self.get_data().boolVal, 0, 0)?
                } else if vt == VT_CY.0 {
                    VarBstrFromCy(self.get_data().cyVal, 0, 0)?
                } else if vt == VT_DATE.0 {
                    VarBstrFromDate(self.get_data().date, 0, 0)?
                } else if vt == VT_DECIMAL.0 {
                    VarBstrFromDec(self.get_data().pdecVal, 0, 0)?
                } else if vt == VT_DISPATCH.0 {
                    if let Some(ref disp) = *self.get_data().pdispVal {
                        VarBstrFromDisp(disp, 0, 0)?
                    } else {
                        BSTR::default()
                    }
                } else if vt == VT_I1.0 {
                    VarBstrFromI1(self.get_data().cVal, 0, 0)?
                } else if vt == VT_I2.0 {
                    VarBstrFromI2(self.get_data().iVal, 0, 0)?
                } else if vt == VT_I4.0 || vt == VT_INT.0 {
                    VarBstrFromI4(self.get_data().lVal, 0, 0)?
                } else if vt == VT_I8.0 {
                    VarBstrFromI8(self.get_data().llVal, 0, 0)?
                } else if vt == VT_R4.0 {
                    VarBstrFromR4(self.get_data().fltVal, 0, 0)?
                } else if vt == VT_R8.0 {
                    VarBstrFromR8(self.get_data().dblVal, 0, 0)?
                } else if vt == VT_UI1.0 {
                    VarBstrFromUI1(self.get_data().bVal, 0, 0)?
                } else if vt == VT_UI2.0 {
                    VarBstrFromUI2(self.get_data().uiVal, 0, 0)?
                } else if vt == VT_UI4.0 || vt == VT_UINT.0 {
                    VarBstrFromUI4(self.get_data().ulVal, 0, 0)?
                } else if vt == VT_UI8.0 {
                    VarBstrFromUI8(self.get_data().ullVal, 0, 0)?
                } else {
                    return Err(Error::new(ERR_TYPE, "Error Variant Type"));
                }
            };
            Ok(str.to_string())
        }
    }
}

impl TryInto<String> for Variant {
    type Error = Error;

    fn try_into(self) -> Result<String> {
        (&self).try_into()
    }
}

/// A Wrapper for windows `SAFEARRAY`
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SafeArray {
    array: SAFEARRAY
}

impl SafeArray {
    pub fn new_vector(vt: VARENUM, len: u32) -> Result<Self> {
        unsafe {
            let array = SafeArrayCreateVector(vt.0 as _, 0, len);

            Ok(Self {
                array: *array
            })
        }
    }
}

impl From<SAFEARRAY> for SafeArray {
    fn from(array: SAFEARRAY) -> Self {
        Self {
            array
        }
    }
}

impl Into<SAFEARRAY> for SafeArray {
    fn into(self) -> SAFEARRAY {
        self.array
    }
}

impl AsRef<SAFEARRAY> for SafeArray {
    fn as_ref(&self) -> &SAFEARRAY {
        &self.array
    }
}

impl Display for SafeArray {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        todo!()
    }
}

#[cfg(test)]
mod tests {
    use windows::Win32::System::Ole::VT_BOOL;

    use crate::variants::Value;
    use crate::variants::Variant;

    #[test]
    fn test_variant_null() {
        let v = Variant::from(Value::NULL);
        assert!(v.is_null());
    }

    #[test]
    fn test_variant_bool() {
        let v: Variant = true.into();
        assert!(v.get_type() == VT_BOOL);

        let b: bool = v.try_into().unwrap();
        assert!(b);

        let val = Variant::from(Value::STRING("true".into()));
        let b_val: bool = val.try_into().unwrap();
        assert!(b_val);
    }

    #[test]
    fn test_variant_string() {
        let s = Variant::from(Value::STRING("Hello".into()));
        assert!(s.is_string());
        assert!(s.get_string().unwrap() == "Hello");
    }
}
