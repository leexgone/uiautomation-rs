use std::fmt::Debug;

use crate::core::UIElement;
use crate::errors::Result;

/// `Condition` is a element filter that can be used in `UIMatcher`.
pub trait Condition {
    fn judge(&self, element: &UIElement) -> Result<bool>;
}

pub struct AndCondition {
    pub left: Box<dyn Condition>,
    pub right: Box<dyn Condition>
}

impl AndCondition {
    pub fn new(left: Box<dyn Condition>, right: Box<dyn Condition>) -> Self {
        Self {
            left,
            right
        }
    }
}

impl Condition for AndCondition {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        let ret = self.left.judge(element)? && self.right.judge(element)?;

        Ok(ret)
    }
}

pub struct OrCondition {
    pub left: Box<dyn Condition>,
    pub right: Box<dyn Condition>
}

impl OrCondition {
    pub fn new(left: Box<dyn Condition>, right: Box<dyn Condition>) -> Self {
        Self {
            left,
            right
        }
    }
}

impl Condition for OrCondition {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        let ret = self.left.judge(element)? || self.right.judge(element)?;
        Ok(ret)
    }
}

#[derive(Debug, Default)]
pub struct NameCondition {
    pub value: String,
    pub casesensitive: bool,
    pub partial: bool
}

// impl Default for NameCondition {
//     fn default() -> Self {
//         Self { 
//             value: Default::default(), 
//             casesensitive: false, 
//             partial: true
//         }
//     }
// }

impl Condition for NameCondition {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        let element_name = element.get_name()?;
        let element_name = element_name.as_str();
        let condition_name = self.value.as_str();

        Ok(
            if self.partial {
                if self.casesensitive {
                    element_name.contains(condition_name)
                } else {
                    let element_name = element_name.to_lowercase();
                    let condition_name = condition_name.to_lowercase();

                    element_name.contains(&condition_name)
                }
            } else {
                if self.casesensitive {
                    element_name == condition_name
                } else {
                    element_name.eq_ignore_ascii_case(condition_name)
                }
            }
        )
    }
}

#[derive(Debug, Default)]
pub struct ClassNameCondition {
    pub classname: String
}

// impl Default for ClassNameCondition {
//     fn default() -> Self {
//         Self { 
//             classname: Default::default() 
//         }
//     }
// }

impl Condition for ClassNameCondition {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        let cur_classname = element.get_classname()?;
        Ok(self.classname == cur_classname)
    }
}

#[derive(Debug, Default)]
pub struct ControlTypeCondition {
    pub control_type: i32
}

impl Condition for ControlTypeCondition {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        let ctrl_type = element.get_control_type()?;
        Ok(self.control_type == ctrl_type)
    }
}