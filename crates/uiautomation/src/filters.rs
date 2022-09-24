use std::fmt::Debug;

use windows::Win32::UI::Accessibility::UIA_CONTROLTYPE_ID;

use super::core::UIElement;
use super::errors::Result;

/// `MatcherFilter` is an element filter that can be used in `UIMatcher`.
pub trait MatcherFilter {
    fn judge(&self, element: &UIElement) -> Result<bool>;
}

pub struct AndFilter {
    pub left: Box<dyn MatcherFilter>,
    pub right: Box<dyn MatcherFilter>
}

impl AndFilter {
    pub fn new(left: Box<dyn MatcherFilter>, right: Box<dyn MatcherFilter>) -> Self {
        Self {
            left,
            right
        }
    }
}

impl MatcherFilter for AndFilter {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        let ret = self.left.judge(element)? && self.right.judge(element)?;

        Ok(ret)
    }
}

pub struct OrFilter {
    pub left: Box<dyn MatcherFilter>,
    pub right: Box<dyn MatcherFilter>
}

impl OrFilter {
    pub fn new(left: Box<dyn MatcherFilter>, right: Box<dyn MatcherFilter>) -> Self {
        Self {
            left,
            right
        }
    }
}

impl MatcherFilter for OrFilter {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        let ret = self.left.judge(element)? || self.right.judge(element)?;
        Ok(ret)
    }
}

#[derive(Debug, Default)]
pub struct NameFilter {
    pub value: String,
    pub casesensitive: bool,
    pub partial: bool
}

impl MatcherFilter for NameFilter {
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
pub struct ClassNameFilter {
    pub classname: String
}

impl MatcherFilter for ClassNameFilter {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        let cur_classname = element.get_classname()?;
        Ok(self.classname == cur_classname)
    }
}

#[derive(Debug, Default)]
pub struct ControlTypeFilter {
    pub control_type: UIA_CONTROLTYPE_ID
}

impl MatcherFilter for ControlTypeFilter {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        let ctrl_type = element.get_control_type()?;
        let is_ctrl = element.is_control_element()?;
        Ok(is_ctrl && self.control_type == ctrl_type)
    }
}

pub struct FnFilter<F> where F: Fn(&UIElement) -> Result<bool> {
    pub filter: Box<F>
}

impl<F> MatcherFilter for FnFilter<F> where F: Fn(&UIElement) -> Result<bool> {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        (self.filter)(element)
    }
}