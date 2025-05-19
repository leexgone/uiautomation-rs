#[cfg(feature = "process")]
use std::cell::RefCell;
use std::fmt::Debug;

// use windows::Win32::Foundation::CloseHandle;
// use windows::Win32::System::Diagnostics::ToolHelp::CreateToolhelp32Snapshot;
// use windows::Win32::System::Diagnostics::ToolHelp::Process32First;
// use windows::Win32::System::Diagnostics::ToolHelp::Process32Next;
// use windows::Win32::System::Diagnostics::ToolHelp::PROCESSENTRY32;
// use windows::Win32::System::Diagnostics::ToolHelp::TH32CS_SNAPPROCESS;
#[cfg(feature = "process")]
use windows::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, Process32First, Process32Next, PROCESSENTRY32, TH32CS_SNAPPROCESS
};

// use crate::controls::ControlType;

use super::types::ControlType;

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

#[derive(Debug)]
pub struct ControlTypeFilter {
    pub control_type: ControlType
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

#[cfg(feature = "process")]
#[derive(Debug)]
pub struct ProcessIdFilter {
    pub pid: u32,
    pub sub_progress: bool,
    progresses: RefCell<Option<Vec<u32>>>
}

#[cfg(feature = "process")]
impl Default for ProcessIdFilter {
    fn default() -> Self {
        Self { 
            pid: Default::default(), 
            sub_progress: true, 
            progresses: Default::default() 
        }
    }
}

#[cfg(feature = "process")]
impl ProcessIdFilter {
    pub fn new(pid: u32, sub_progress: bool) -> Self {
        Self {
            pid,
            sub_progress,
            progresses: RefCell::new(None)
        }
    }

    fn contains(&self, pid: u32) -> Result<bool> {
        if !self.is_valid() {
            self.build_sub_ids()?;
        }

        let ids = self.progresses.borrow();
        if let Some(ids) = ids.as_ref() {
            Ok(ids.contains(&pid))
        } else {
            Ok(false)
        }
    }

    fn is_valid(&self) -> bool {
        if let Some(ids) = self.progresses.borrow().as_ref() {
            ids.get(0) == Some(&self.pid)
        } else {
            false
        }
    }

    fn build_sub_ids(&self) -> Result<()> {
        let procs = self.get_sub_processes()?;

        let mut sub_ids = vec![self.pid];
        let mut found = true;
        while found {
            found = false;
            for (pid, ppid) in procs.iter() {
                if sub_ids.contains(ppid) && !sub_ids.contains(pid) {
                    sub_ids.push(*pid);
                    found = true;
                }
            }
        }

        let mut ids = self.progresses.borrow_mut();
        *ids = Some(sub_ids);

        Ok(())
    }

    fn get_sub_processes(&self) -> Result<Vec<(u32, u32)>> {
        let mut ids = Vec::new();

        unsafe  {
            let snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)?;

            let mut proc_entry = PROCESSENTRY32::default();
            proc_entry.dwSize = std::mem::size_of::<PROCESSENTRY32>() as _;

            let mut found = Process32First(snapshot, &mut proc_entry);
            while found.is_ok() {
                ids.push((proc_entry.th32ProcessID, proc_entry.th32ParentProcessID));

                found = Process32Next(snapshot, &mut proc_entry);
            }

            windows::Win32::Foundation::CloseHandle(snapshot)?;
        }

        Ok(ids)
    }
}

#[cfg(feature = "process")]
impl MatcherFilter for ProcessIdFilter {
    fn judge(&self, element: &UIElement) -> Result<bool> {
        let pid = element.get_process_id()?;
        if self.pid == pid {
            Ok(true)
        } else if self.sub_progress {
            self.contains(pid)
        } else {
            Ok(false)
        }
    }
}