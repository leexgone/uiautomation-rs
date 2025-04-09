use std::mem;
use std::iter::once;

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
use windows::Win32::System::Threading::WaitForInputIdle;
use windows::Win32::System::Threading::WaitForSingleObject;
use windows::core::PCWSTR;
use windows::core::PWSTR;

use super::Error;
use super::Result;
use crate::errors::ERR_ALREADY_RUNNING;
use crate::types::Handle;
use super::errors::ERR_NONE;
use super::errors::ERR_TIMEOUT;

/// Windows process wrapper.
#[derive(Debug)]
pub struct Process {
    application: Option<String>,
    command: Option<String>,
    cur_dir: Option<String>,
    wait_for_idle: Option<u32>,
    startup_info: STARTUPINFOW,
    proc_info: PROCESS_INFORMATION
}

struct WSTR {
    data: Option<Vec<u16>>,
}

impl WSTR {
    fn new(s: Option<&str>) -> Self {
        Self {
            data: s.map(|s| s.encode_utf16().chain(once(0)).collect()),
        }
    }

    fn to_pcwstr(&self) -> PCWSTR {
        self.data
            .as_ref()
            .map(|s| PCWSTR::from_raw(s.as_ptr()))
            .unwrap_or_else(|| PCWSTR::null())
    }

    fn to_pwstr(&mut self) -> PWSTR {
        self.data
            .as_mut()
            .map(|s| PWSTR::from_raw(s.as_mut_ptr()))
            .unwrap_or_else(|| PWSTR::null())
    }
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
            application: None,
            command: Some(command.into()),
            cur_dir: None,
            wait_for_idle: None,
            startup_info: Self::startupinfo(), 
            proc_info: PROCESS_INFORMATION::default(),
        }
    }

    /// Set the name of the module to be executed.
    pub fn application<S: Into<String>>(mut self, application: S) -> Self {
        self.application = Some(application.into());
        self
    }

    /// Set the command line to be executed.
    pub fn command<S: Into<String>>(mut self, command: S) -> Self {
        self.command = Some(command.into());
        self
    }

    /// Set the current directory as `dir`, which is the full path to the current directory for the process. 
    pub fn current_directory<S: Into<String>>(mut self, dir: S) -> Self {
        self.cur_dir = Some(dir.into());
        self
    }

    /// Set to wait until the specified process has finished processing its initial input and is waiting for user input with no input pending, 
    /// or until the time-out interval(`milliseconds`) has elapsed.
    pub fn wait_for_idle(mut self, milliseconds: u32) -> Self {
        self.wait_for_idle = Some(milliseconds);
        self
    }

    /// Run the current process.
    pub fn run(&mut self) -> Result<()> {
        if !self.proc_info.hProcess.is_invalid() {
            Err(Error::new(ERR_ALREADY_RUNNING, "process is already started"))
        } else {
            let app = WSTR::new(self.application.as_deref());
            let mut cmd = WSTR::new(self.command.as_deref());
            let cur_dir = WSTR::new(self.cur_dir.as_deref());

            unsafe {
                CreateProcessW(
                    app.to_pcwstr(),
                    Some(cmd.to_pwstr()),
                    None,
                    None,
                    true,
                    PROCESS_CREATION_FLAGS::default(),
                    None,
                    cur_dir.to_pcwstr(),
                    &self.startup_info,
                    &mut self.proc_info)?
            };

            if let Some(timeout) = self.wait_for_idle {
                unsafe { WaitForInputIdle(self.proc_info.hProcess, timeout) };
            }
    
            Ok(())
        }
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

    /// Get the process handle.
    pub fn get_handle(&self) -> Handle {
        let h = windows::Win32::Foundation::HWND(self.proc_info.hProcess.0);
        h.into()
    }

    /// Get the process ID.
    pub fn get_id(&self) -> u32 {
        self.proc_info.dwProcessId
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

impl Default for Process {
    fn default() -> Self {
        Self { 
            application: None, 
            command: None, 
            cur_dir: None,
            wait_for_idle: None,
            startup_info: Self::startupinfo(), 
            proc_info: Default::default() 
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::processes::Process;

    #[test]
    fn run_notepad() {
        let proc_notepad = Process::create("notepad.exe");
        assert!(proc_notepad.is_ok());
    }

    #[test]
    fn run_ping() {
        let mut ping = Process::new("ping.exe localhost -n 1").current_directory("C:/");
        ping.run().unwrap();
        ping.wait().unwrap();
    }

    #[test]
    fn run_calc() {
        let mut calc = Process::default()
            .application("C:\\Windows\\System32\\calc.exe")
            .current_directory("C:\\")
            .wait_for_idle(5000);
        calc.run().unwrap();
    }
}