use std::mem;

use windows::Win32::Foundation::CloseHandle;
use windows::Win32::Foundation::HANDLE;
use windows::Win32::Foundation::WAIT_FAILED;
use windows::Win32::Foundation::WAIT_OBJECT_0;
use windows::Win32::Foundation::WAIT_TIMEOUT;
use windows::Win32::System::Threading::CreateProcessW;
use windows::Win32::System::Threading::GetExitCodeProcess;
use windows::Win32::System::Threading::INFINITE;
use windows::Win32::System::Threading::PROCESS_CREATION_FLAGS;
use windows::Win32::System::Threading::PROCESS_INFORMATION;
use windows::Win32::System::Threading::STARTUPINFOW;
use windows::Win32::System::Threading::TerminateProcess;
use windows::Win32::System::Threading::WaitForSingleObject;
use windows::core::PCWSTR;
use windows::core::PWSTR;

use super::Error;
use super::Result;
use crate::errors::ERR_ALREADY_RUNNING;
use super::errors::ERR_NONE;
use super::errors::ERR_TIMEOUT;

/// Windows process wrapper.
#[derive(Debug)]
pub struct Process {
    command: String,
    cur_dir: Option<String>,
    startup_info: STARTUPINFOW,
    proc_info: PROCESS_INFORMATION
}

impl Process {
    /// Create and run a process by command line.
    /// 
    /// # Examples
    /// ```
    /// use uiautomation::processes::Process;
    /// 
    /// let p = Process::create("notepad.exe");
    /// assert!(p.is_ok());
    /// ```
    pub fn create(command: &str) -> Result<Self> {
        let mut process = Self::new(command);
        process.run()?;
        Ok(process)
    }

    #[inline]
    fn startupinfo() -> STARTUPINFOW {
        let mut si = STARTUPINFOW::default();
        si.cb = mem::size_of::<STARTUPINFOW>() as _;
        si
    }

    /// Create a process by `command`. This process will not startup until it is called by `run()`. 
    pub fn new<S: Into<String>>(command: S) -> Self {
        Self { 
            command: command.into(),
            cur_dir: None,
            startup_info: Self::startupinfo(), 
            proc_info: PROCESS_INFORMATION::default(),
        }
    }

    /// Set the current directory as `dir`, which is the full path to the current directory for the process. 
    pub fn current_directory<S: Into<String>>(mut self, dir: S) -> Self {
        self.cur_dir = Some(dir.into());
        self
    }

    /// Run the current process.
    pub fn run(&mut self) -> Result<()> {
        if !self.proc_info.hProcess.is_invalid() {
            Err(Error::new(ERR_ALREADY_RUNNING, "process is already started"))
        } else {
            let cmd = {
                let mut buf: Vec<u16> = self.command.encode_utf16().chain(std::iter::once(0)).collect();
                PWSTR::from_raw(buf.as_mut_ptr())
            };
            let cur_dir = if let Some(ref val) = self.cur_dir {
                let buf: Vec<u16> = val.encode_utf16().chain(std::iter::once(0)).collect();
                PCWSTR::from_raw(buf.as_ptr())
            } else {
                PCWSTR::null()
            };

            unsafe {
                CreateProcessW(PCWSTR::null(), 
                    cmd,
                    None, 
                    None, 
                    true, 
                    PROCESS_CREATION_FLAGS::default(), 
                    None,
                    cur_dir,

                    &self.startup_info,
                    &mut self.proc_info)?
            };
    
            Ok(())
        }
    }

    /// Exit the process with `exit_code` by force.
    pub fn terminate(&self, exit_code: u32) -> Result<()> {
        unsafe {
            TerminateProcess(self.proc_info.hProcess, exit_code)?
        };
        Ok(())
    }

    /// Wait for the process to exit.
    /// 
    /// `timeout` is the milliseconds to wait for.
    pub fn wait_for(&self, timeout: u32) -> Result<()> {
        let ret = unsafe {
            WaitForSingleObject(self.proc_info.hProcess, timeout)
        };

        if ret == WAIT_OBJECT_0 {
            Ok(())
        } else if ret == WAIT_FAILED {
            Err(Error::last_os_error())
        } else if ret == WAIT_TIMEOUT {
            Err(Error::new(ERR_TIMEOUT, "Wait Timeout"))
        } else {
            Err(Error::new(ERR_NONE, "Wait Failed"))
        }
    }

    /// Wait until the process exits.
    pub fn wait(&self) -> Result<()> {
        self.wait_for(INFINITE)
    }

    /// Get the exit code of the process.
    pub fn get_exit_code(&self) -> Result<u32> {
        let mut exit_code: u32 = 0;
        unsafe {
            GetExitCodeProcess(self.proc_info.hProcess, &mut exit_code)?
        };

        Ok(exit_code)
    }
}

macro_rules! close_handle {
    ($handle: expr) => {
        if !$handle.is_invalid() {
            let _ = unsafe { CloseHandle($handle) };
            $handle = HANDLE::default();
        }
    };
}

impl Drop for Process {
    fn drop(&mut self) {
        close_handle!(self.startup_info.hStdInput);
        close_handle!(self.startup_info.hStdOutput);
        close_handle!(self.startup_info.hStdError);

        close_handle!(self.proc_info.hThread);
        close_handle!(self.proc_info.hProcess);
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

    #[test]
    fn run_ping() {
        let mut ping = Process::new("ping.exe localhost -n 1").current_directory("C:/");
        ping.run().unwrap();
        ping.wait().unwrap();
    }
}