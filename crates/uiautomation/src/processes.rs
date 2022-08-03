use std::mem;

use windows::Win32::Foundation::CloseHandle;
use windows::Win32::Foundation::WAIT_FAILED;
use windows::Win32::Foundation::WAIT_OBJECT_0;
use windows::Win32::Foundation::WAIT_TIMEOUT;
use windows::Win32::System::Threading::CreateProcessW;
use windows::Win32::System::Threading::GetExitCodeProcess;
use windows::Win32::System::Threading::PROCESS_CREATION_FLAGS;
use windows::Win32::System::Threading::PROCESS_INFORMATION;
use windows::Win32::System::Threading::STARTUPINFOW;
use windows::Win32::System::Threading::TerminateProcess;
use windows::Win32::System::Threading::WaitForSingleObject;
use windows::core::PCWSTR;
use windows::core::PWSTR;

use super::Error;
use super::Result;
use super::errors::ERR_NONE;
use super::errors::ERR_TIMEOUT;

/// Windows process wrapper.
#[derive(Debug)]
pub struct Process {
    information: PROCESS_INFORMATION
}

impl Process {
    /// Create process by command line.
    /// 
    /// # Examples
    /// ```
    /// use uiautomation::processes::Process;
    /// 
    /// let p = Process::create("notepad.exe");
    /// assert!(p.is_ok());
    /// ```
    pub fn create(command: &str) -> Result<Self> {
        let mut information = PROCESS_INFORMATION::default();
        let mut buffer = command.encode_utf16().chain(std::iter::once(0)).collect::<Vec<u16>>();
        let si = Process::startupinfo();
        let ret = unsafe {
            CreateProcessW(PCWSTR::null(), 
                PWSTR(buffer.as_mut_ptr()), 
                std::ptr::null(), 
                std::ptr::null(), 
                false, 
                PROCESS_CREATION_FLAGS::default(), 
                std::ptr::null(),
                PCWSTR::null(),
                &si,
                &mut information)
        };

        if ret.as_bool() {
            Ok(Self {
                information
            })
        } else {
            Err(Error::last_os_error())
        }
    }

    #[inline]
    fn startupinfo() -> STARTUPINFOW {
        let mut si = STARTUPINFOW::default();
        si.cb = mem::size_of::<STARTUPINFOW>() as _;
        si
    }

    /// Exit the process with `exit_code` by force.
    pub fn terminate(&self, exit_code: u32) -> Result<()> {
        let ret = unsafe {
            TerminateProcess(self.information.hProcess, exit_code)
        };

        if ret.as_bool() {
            Ok(())
        } else {
            Err(Error::last_os_error())
        }
    }

    /// Wait for the process to exit.
    /// 
    /// `timeout` is the milliseconds to wait for.
    pub fn wait_for(&self, timeout: u32) -> Result<()> {
        let ret = unsafe {
            WaitForSingleObject(self.information.hProcess, timeout)
        };

        if ret == WAIT_OBJECT_0.0 {
            Ok(())
        } else if ret == WAIT_FAILED.0 {
            Err(Error::last_os_error())
        } else if ret == WAIT_TIMEOUT.0 {
            Err(Error::new(ERR_TIMEOUT, "Wait Timeout"))
        } else {
            Err(Error::new(ERR_NONE, "Wait Failed"))
        }
    }

    /// Get the exit code of the process.
    pub fn get_exit_code(&self) -> Result<u32> {
        let mut exit_code: u32 = 0;
        let ret = unsafe {
            GetExitCodeProcess(self.information.hProcess, &mut exit_code)
        };

        if ret.as_bool() {
            Ok(exit_code)
        } else {
            Err(Error::last_os_error())
        }
    }
}

impl Drop for Process {
    fn drop(&mut self) {
        unsafe {
            CloseHandle(self.information.hProcess);
            CloseHandle(self.information.hThread);
        }
    }
}

#[cfg(test)]
mod tests {
    use std::thread::sleep;
    use std::time::Duration;

    use crate::processes::Process;

    #[test]
    fn run_notepad() {
        let proc_notepad = Process::create("notepad.exe");
        assert!(proc_notepad.is_ok());

        sleep(Duration::from_secs(1));

        let notepad = proc_notepad.unwrap();
        let ret = notepad.terminate(1);
        assert!(ret.is_ok());

        let exit_code = notepad.get_exit_code().unwrap();
        assert_eq!(exit_code, 1);
    }
}