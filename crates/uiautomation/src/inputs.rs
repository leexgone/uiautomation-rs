use std::cmp::max;
use std::collections::HashMap;
use std::collections::HashSet;
use std::mem;
use std::str::Chars;
use std::sync::OnceLock;
use std::thread::sleep;
use std::time::Duration;

use windows::Win32::UI::Input::KeyboardAndMouse::*;
use windows::Win32::UI::WindowsAndMessaging::GetCursorPos;
use windows::Win32::UI::WindowsAndMessaging::GetSystemMetrics;
use windows::Win32::UI::WindowsAndMessaging::SM_CXSCREEN;
use windows::Win32::UI::WindowsAndMessaging::SM_CYSCREEN;
use windows::Win32::UI::WindowsAndMessaging::SetCursorPos;

use crate::log_error;

use super::errors::{ERR_FORMAT, ERR_INVALID_ARG};
use super::Error;
use super::Result;
use super::types::Point;

const KEYEVENTF_KEYDOWN: KEYBD_EVENT_FLAGS = KEYBD_EVENT_FLAGS(0);

macro_rules! map {
    ( $($key: expr_2021 => $value: expr_2021), * ) => {
        {
            let mut map = std::collections::HashMap::new();
            $(
                map.insert($key.to_string(), $value);
            )*
            map
        }
    };
}

macro_rules! set {
    ( $($value: expr_2021), * ) => {
        {
            let mut set = std::collections::HashSet::new();
            $(
                set.insert($value.to_string());
            )*
            set
        }
    };
}

static VIRTUAL_KEYS: OnceLock<HashMap<String, VIRTUAL_KEY>> = OnceLock::new();
static HOLD_KEYS: OnceLock<HashSet<String>> = OnceLock::new();

fn init_virtual_keys() -> HashMap<String, VIRTUAL_KEY> {
    map![
        "CONTROL" => VK_CONTROL, "CTRL" => VK_CONTROL, "LCONTROL" => VK_LCONTROL, "LCTRL" => VK_LCONTROL, "CTRLLEFT" => VK_LCONTROL, "RCONTROL" => VK_RCONTROL, "RCTRL" => VK_RCONTROL, "CTRLRIGHT" => VK_RCONTROL,
        "ALT" => VK_MENU, "MENU" => VK_MENU, "LALT" => VK_LMENU, "LMENU" => VK_LMENU, "ALTLEFT" => VK_LMENU, "RALT" => VK_RMENU, "RMENU" => VK_RMENU, "ALTRIGHT" => VK_RMENU,
        "SHIFT" => VK_SHIFT, "LSHIFT" => VK_LSHIFT, "SHIFTLEFT" => VK_LSHIFT, "RSHIFT" => VK_RSHIFT, "SHIFTRIGHT" => VK_RSHIFT,
        "WIN" => VK_LWIN, "WINDOWS" => VK_LWIN, "LWIN" => VK_LWIN, "LWINDOWS" => VK_LWIN, "WINLEFT" => VK_LWIN, "RWIN" => VK_RWIN, "RWINDOWS" => VK_RWIN, "WINRIGHT" => VK_RWIN,
        "LBUTTON" => VK_LBUTTON, "RBUTTON" => VK_RBUTTON, "MBUTTON" => VK_MBUTTON, "XBUTTON1" => VK_XBUTTON1, "XBUTTON2" => VK_XBUTTON2,
        "CANCEL" => VK_CANCEL, "BACKSPACE" => VK_BACK, "BACK" => VK_BACK, "TAB" => VK_TAB, "RETURN" => VK_RETURN, "ENTER" => VK_RETURN, "PAUSE" => VK_PAUSE, "CAPITAL" => VK_CAPITAL, "CAPSLOCK" => VK_CAPITAL,
        "ESCAPE" => VK_ESCAPE, "ESC" => VK_ESCAPE, "SPACE" => VK_SPACE,
        "PRIOR" => VK_PRIOR, "PAGE_UP" => VK_PRIOR, "PAGEUP" => VK_PRIOR, "PGUP" => VK_PRIOR, "NEXT" => VK_NEXT, "PAGE_DOWN" => VK_NEXT, "PAGEDOWN" => VK_NEXT, "PGDN" => VK_NEXT,
        "HOME" => VK_HOME, "END" => VK_END,
        "LEFT" => VK_LEFT, "UP" => VK_UP, "RIGHT" => VK_RIGHT, "DOWN" => VK_DOWN,
        "PRINT" => VK_PRINT, "PRINTSCREEN" => VK_PRINT, "PRNTSCRN" => VK_PRINT, "PRTSC" => VK_PRINT, "PRTSCR" => VK_PRINT,
        "INSERT" => VK_INSERT, "INS" => VK_INSERT, "DELETE" => VK_DELETE, "DEL" => VK_DELETE,
        "F1" => VK_F1, "F2" => VK_F2, "F3" => VK_F3, "F4" => VK_F4, "F5" => VK_F5, "F6" => VK_F6, "F7" => VK_F7, "F8" => VK_F8, "F9" => VK_F9, "F10" => VK_F10,
        "F11" => VK_F11, "F12" => VK_F12, "F13" => VK_F13, "F14" => VK_F14, "F15" => VK_F15, "F16" => VK_F16, "F17" => VK_F17, "F18" => VK_F18, "F19" => VK_F19,
        "F20" => VK_F20, "F21" => VK_F21, "F22" => VK_F22, "F23" => VK_F23, "F24" => VK_F24
    ]
}

fn init_hold_keys() -> HashSet<String> {
    set![
        "CONTROL", "CTRL", "LCONTROL", "LCTRL", "RCONTROL", "RCTRL",
        "ALT", "MENU", "LALT", "LMENU", "RALT", "RMENU",
        "SHIFT", "LSHIFT", "RSHIFT",
        "WIN", "WINDOWS", "LWIN", "LWINDOWS", "RWIN", "RWINDOWS"
    ]
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InputItem {
    HoldKey(VIRTUAL_KEY),
    VirtualKey(VIRTUAL_KEY, usize), // usize is the repeat count.
    Character(char),
}

impl InputItem {
    pub fn is_holdkey(&self) -> bool {
        match self {
            Self::HoldKey(_) => true,
            _ => false,
        }
    }
}
/// Represents an input item, which can be a hold key, a virtual key, or a character.
#[derive(Debug, PartialEq, Eq, Default)]
pub struct Input {
    holdkeys: Vec<VIRTUAL_KEY>,
    items: Vec<InputItem>,
}

impl Input {
    pub fn new() -> Self {
        Self {
            holdkeys: Vec::new(),
            items: Vec::new(),
        }
    }

    pub fn has_holdkey(&self) -> bool {
        !self.holdkeys.is_empty()
    }

    pub fn has_items(&self) -> bool {
        !self.items.is_empty()
    }

    pub fn is_holdkey_only(&self) -> bool {
        !self.holdkeys.is_empty() && self.items.is_empty()
    }

    pub fn push(&mut self, item: InputItem) {
        if let InputItem::HoldKey(key) = item {
            if !self.holdkeys.contains(&key) {
                self.holdkeys.push(key);
            }
        } else {
            self.items.push(item);
        }
    }

    pub fn push_all(&mut self, items: &Vec<InputItem>) {
        for item in items {
            self.push(*item);
        }
    }

    /// Create multiple input items to the current input.
    pub fn create_inputs(&self) -> Result<Vec<INPUT>> {
        self.create_inputs_internal(false)
    }

    /// like create_inputs but always create unicode inputs. This is more robust but might not simulate key presses like from a real keyboard.
    pub fn create_unicode_inputs(&self) -> Result<Vec<INPUT>> {
        self.create_inputs_internal(true)
    }

    fn create_inputs_internal(&self, force_unicode: bool) -> Result<Vec<INPUT>> {
        let mut inputs: Vec<INPUT> = Vec::new();

        for holdkey in &self.holdkeys {
            let input = Self::create_virtual_key(*holdkey, KEYEVENTF_KEYDOWN);
            inputs.push(input);
        }

        for item in &self.items {
            match item {
                InputItem::VirtualKey(key, count) => {
                    for _ in 0..max(*count, 1) {
                        inputs.push(Self::create_virtual_key(*key, KEYEVENTF_KEYDOWN));
                        inputs.push(Self::create_virtual_key(*key, KEYEVENTF_KEYUP));
                    }
                },
                InputItem::Character(ch) => {
                    let mut buffer = [0; 2];
                    let chars = ch.encode_utf16(&mut buffer);
                    for ch_u16 in chars {
                        let keys = Self::create_char_key(*ch_u16, self.has_holdkey(), force_unicode);
                        inputs.extend(keys);
                    }
                },
                _ => (),
            }
        }

        for holdkey in (&self.holdkeys).iter().rev() {
            let input = Self::create_virtual_key(*holdkey, KEYEVENTF_KEYUP);
            inputs.push(input);
        }

        Ok(inputs)
    }

    fn create_virtual_key(key: VIRTUAL_KEY, flags: KEYBD_EVENT_FLAGS) -> INPUT {
        INPUT {
            r#type: INPUT_KEYBOARD,
            Anonymous: INPUT_0 {
                ki: KEYBDINPUT {
                    wVk: key,
                    wScan: 0,
                    dwFlags: flags,
                    time: 0,
                    dwExtraInfo: 0,
                },
            },
        }
    }

    fn create_char_key(ch: u16, hold_mode: bool, force_unicode: bool) -> Vec<INPUT> {
        // let code = ch as i32;
        let vk: i16 = if !force_unicode && (ch < 256) {
            unsafe { VkKeyScanW(ch) }
        } else {
            -1
        };
        
        if vk == -1 {   // Unicode
            vec![
                INPUT {
                    r#type: INPUT_KEYBOARD,
                    Anonymous: INPUT_0 {
                        ki: KEYBDINPUT {
                            wVk: VIRTUAL_KEY(0),
                            wScan: ch,
                            dwFlags: KEYEVENTF_UNICODE,
                            time: 0,
                            dwExtraInfo: 0,
                            },
                        }
                }, INPUT {
                    r#type: INPUT_KEYBOARD,
                    Anonymous: INPUT_0 {
                        ki: KEYBDINPUT {
                            wVk: VIRTUAL_KEY(0),
                            wScan: ch,
                            dwFlags: KEYEVENTF_UNICODE | KEYEVENTF_KEYUP,
                            time: 0,
                            dwExtraInfo: 0,
                        },
                    }
                }
            ]
        } else {        // ASCII
            let key: VIRTUAL_KEY = VIRTUAL_KEY((vk & 0xFF) as _);
            if hold_mode {
                vec![
                    INPUT {
                        r#type: INPUT_KEYBOARD,
                        Anonymous: INPUT_0 {
                            ki: KEYBDINPUT {
                                wVk: key,
                                wScan: 0,
                                dwFlags: KEYEVENTF_KEYDOWN,
                                time: 0,
                                dwExtraInfo: 0,
                                },
                            }
                    }, INPUT {
                        r#type: INPUT_KEYBOARD,
                        Anonymous: INPUT_0 {
                            ki: KEYBDINPUT {
                                wVk: key,
                                wScan: 0,
                                dwFlags: KEYEVENTF_KEYUP,
                                time: 0,
                                dwExtraInfo: 0,
                            },
                        }
                    }
                ]    
            } else {
                let mut shift: bool = (vk & 0x0100) != 0;
                let ctrl: bool = (vk & 0x0200) != 0;
                let alt: bool = (vk & 0x0400) != 0;
                // Do not press shift if CAPSLOCK is on.
                let state = unsafe { GetKeyState(VK_CAPITAL.0 as _) };
                if (state & 0x01) != 0 {
                    if (ch >= 'a' as u16 && ch <= 'z' as u16) || (ch >= 'A' as u16 && ch <= 'Z' as u16) {
                        shift = !shift;
                    }
                };
                let mut char_inputs: Vec<INPUT> = Vec::new();
                if shift {
                    char_inputs.push(
                        INPUT {
                            r#type: INPUT_KEYBOARD,
                            Anonymous: INPUT_0 {
                                ki: KEYBDINPUT {
                                    wVk: VK_SHIFT,
                                    wScan: 0,
                                    dwFlags: KEYEVENTF_KEYDOWN,
                                    time: 0,
                                    dwExtraInfo: 0,
                                    },
                                }
                        }
                    );
                }
                if ctrl {
                    char_inputs.push(
                        INPUT {
                            r#type: INPUT_KEYBOARD,
                            Anonymous: INPUT_0 {
                                ki: KEYBDINPUT {
                                    wVk: VK_CONTROL,
                                    wScan: 0,
                                    dwFlags: KEYEVENTF_KEYDOWN,
                                    time: 0,
                                    dwExtraInfo: 0,
                                    },
                                }
                        }
                    );
                }
                if alt {
                    char_inputs.push(
                        INPUT {
                            r#type: INPUT_KEYBOARD,
                            Anonymous: INPUT_0 {
                                ki: KEYBDINPUT {
                                    wVk: VK_MENU,
                                    wScan: 0,
                                    dwFlags: KEYEVENTF_KEYDOWN,
                                    time: 0,
                                    dwExtraInfo: 0,
                                    },
                                }
                        }
                    );
                }
                char_inputs.push(
                    INPUT {
                        r#type: INPUT_KEYBOARD,
                        Anonymous: INPUT_0 {
                            ki: KEYBDINPUT {
                                wVk: key,
                                wScan: 0,
                                dwFlags: KEYEVENTF_KEYDOWN,
                                time: 0,
                                dwExtraInfo: 0,
                                },
                            }
                    }
                );
                char_inputs.push(INPUT {
                        r#type: INPUT_KEYBOARD,
                        Anonymous: INPUT_0 {
                            ki: KEYBDINPUT {
                                wVk: key,
                                wScan: 0,
                                dwFlags: KEYEVENTF_KEYUP,
                                time: 0,
                                dwExtraInfo: 0,
                            },
                        }
                    }
                );
                if alt {
                    char_inputs.push(
                        INPUT {
                            r#type: INPUT_KEYBOARD,
                            Anonymous: INPUT_0 {
                                ki: KEYBDINPUT {
                                    wVk: VK_MENU,
                                    wScan: 0,
                                    dwFlags: KEYEVENTF_KEYUP,
                                    time: 0,
                                    dwExtraInfo: 0,
                                },
                            }
                        }
                    );
                }
                if ctrl {
                    char_inputs.push(
                        INPUT {
                            r#type: INPUT_KEYBOARD,
                            Anonymous: INPUT_0 {
                                ki: KEYBDINPUT {
                                    wVk: VK_CONTROL,
                                    wScan: 0,
                                    dwFlags: KEYEVENTF_KEYUP,
                                    time: 0,
                                    dwExtraInfo: 0,
                                },
                            }
                        }
                    );
                }
                if shift {
                    char_inputs.push(
                        INPUT {
                            r#type: INPUT_KEYBOARD,
                            Anonymous: INPUT_0 {
                                ki: KEYBDINPUT {
                                    wVk: VK_SHIFT,
                                    wScan: 0,
                                    dwFlags: KEYEVENTF_KEYUP,
                                    time: 0,
                                    dwExtraInfo: 0,
                                    },
                                }
                        }
                    );
                };
                char_inputs
            }
        }
    }
}

#[derive(Debug)]
pub struct Parser {
    ignore_err: bool,
}

impl Parser {
    pub fn new(ignore_err: bool) -> Self {
        Self { ignore_err }
    }

    pub fn parse_input(&self, expression: &str) -> Result<Vec<Input>> {
        let mut inputs: Vec<Input> = Vec::new();

        let mut expr = expression.chars();
        while let Some((items, is_holdkey)) = self.next_input(&mut expr, if let Some(p) = inputs.last() { p.is_holdkey_only() } else { false })? {
            if let Some(prev) = inputs.last_mut() {
                if (is_holdkey && !prev.has_items()) || (!is_holdkey && (!prev.has_holdkey() || prev.is_holdkey_only())) { 
                    prev.push_all(&items);
                    continue;
                }
            }

            let mut input = Input::new();
            input.push_all(&items);

            inputs.push(input);
        }

        Ok(inputs)
    }

    fn next_input(&self, expr: &mut Chars<'_>, is_holding: bool) -> Result<Option<(Vec<InputItem>, bool)>> {
        if let Some(ch) = expr.next() {
            let next = match ch {
                '{' => {
                    if self.ignore_err {
                        let snapshot = expr.clone();
                        if let Ok(item) = self.read_special_item(expr) {
                            Some((vec![item], item.is_holdkey()))
                        } else {
                            *expr = snapshot;
                            Some((vec![InputItem::Character(ch)], false))
                        }
                    } else {
                        let item =  self.read_special_item(expr)?;
                        Some((vec![item], item.is_holdkey()))
                    }
                }
                '(' => {
                    if !is_holding {
                        Some((vec![InputItem::Character(ch)], false))
                    } else if self.ignore_err {
                        let snapshot = expr.clone();
                        if let Ok(items) = self.read_group_items(expr) {
                            Some((items, false))
                        } else {
                            *expr = snapshot;
                            Some((vec![InputItem::Character(ch)], false))
                        }
                    } else {
                        let items = self.read_group_items(expr)?;
                        Some((items, false))
                    }
                }
                _ => Some((vec![InputItem::Character(ch)], false)),
            };
            Ok(next)
        } else {
            Ok(None)
        }
    }

    fn read_special_item(&self, expr: &mut Chars<'_>) -> Result<InputItem> {
        let mut token = String::new();
        let mut matched = false;
        while let Some(ch) = expr.next() {
            if ch == '}' && !token.is_empty() {
                matched = true;
                break;
            } else {
                token.push(ch);
            }
        }

        if matched {
            if token == "(" || token == ")" || token == "{" || token == "}" {
                Ok(InputItem::Character(token.chars().nth(0).unwrap()))
            } else {
                let vkeys = VIRTUAL_KEYS.get_or_init(init_virtual_keys); 
                let hkeys = HOLD_KEYS.get_or_init(init_hold_keys);
                
                // let token = token.to_uppercase();
                let mut itr = token.split_whitespace();
                let expr = itr.next().unwrap().to_uppercase();
                let modifier = itr.next();
                if let Some(key) = vkeys.get(&expr) {
                    if hkeys.contains(&expr) {
                        Ok(InputItem::HoldKey(*key))
                    } else {
                        let count: usize = modifier.and_then(|s| s.parse().ok()).unwrap_or(1);
                        Ok(InputItem::VirtualKey(*key, count))
                    }
                } else {
                    Err(Error::new(ERR_FORMAT, "Error Input Format"))
                }
            }
        } else {
            Err(Error::new(ERR_FORMAT, "Error Input Format"))
        }
    }

    fn read_group_items(&self, expr: &mut Chars<'_>) -> Result<Vec<InputItem>> {
        let mut items: Vec<InputItem> = Vec::new();
        let mut matched = false;

        while let Some((next, _)) = self.next_input(expr, if let Some(&InputItem::HoldKey(_)) = items.last() { true } else { false })? {
            if next.len() == 1 && next[0] == InputItem::Character(')') {
                matched = true;
                break;
            }

            items.extend(next);
        }

        if matched && !items.is_empty() {
            Ok(items)
        } else {
            Err(Error::new(ERR_FORMAT, "Error Input Format"))
        }
    }
}

/// Simulate typing keys on keyboard.
#[derive(Debug, Default)]
pub struct Keyboard {
    interval: u64,
    holdkeys: Vec<VIRTUAL_KEY>,
    ignore_parse_err: bool,
}

impl Keyboard {
    /// Create a keyboard to simulate typing keys.
    pub fn new() -> Self {
        Self {
            interval: 50,
            holdkeys: Vec::new(),
            ignore_parse_err: false,
        }
    }

    /// Set the interval time between keys.
    /// 
    /// `interval` is the time number of milliseconds, `50` is default value.
    /// 
    /// When inputting long texts, if the `interval` value is set too low, input loss may occur during the input process.
    pub fn interval(mut self, interval: u64) -> Self {
        self.interval = interval;
        self
    }

    /// Set whether ignore parse error.
    /// 
    /// `ignore_parse_err` is the flag to ignore parse error, `false` is default value.
    /// 
    /// If `ignore_parse_err` is `true`, the parser will ignore parse error and allow `{` & `(` as regular inputs.
    /// For example: `{Hi},(rust)!` types `{Hi},(rust)!`.
    pub fn ignore_parse_err(mut self, ignore_parse_err: bool) -> Self {
        self.ignore_parse_err = ignore_parse_err;
        self
    }

    /// Simulates typing `keys` on keyboard.
    /// 
    /// `{}` is used for some special keys. For example: `{ctrl}{alt}{delete}`, `{shift}{home}`.
    /// 
    /// Special keys can specify the number of repetitions. For example: '{enter 3}' means pressing the Enter key three times.
    /// 
    /// `()` is used for group keys. The '(' symbol only takes effect after the '{}' symbol. For example: `{ctrl}(AB)` types `Ctrl+A+B`.
    /// 
    /// `{` `}` `(` `)` can be quoted by `{}`. For example: `{{}Hi,{(}rust!{)}{}}` types `{Hi,(rust)}`.
    /// 
    ///  When `ignore_parse_err` is `false`, the parser will throw an error if the input format is incorrect.
    ///  When `ignore_parse_err` is `true`, the parser will ignore parse error and try to allow `{` & `(` as regular inputs.
    pub fn send_keys(&self, keys: &str) -> Result<()> {
        let inputs = Parser::new(self.ignore_parse_err).parse_input(keys)?;
        for ref input in inputs {
            // self.send_keyboard(input)?;
            let input_keys = input.create_inputs()?;
            self.send_keyboard(&input_keys)?;
        }

        Ok(())
    }

    /// Simulates typing `text` on keyboard without any special keys.
    /// 
    /// This method will only output the literal content of the text.
    pub fn send_text(&self, text: &str) -> Result<()> {
        let inputs = vec![
            Input {
                holdkeys: Vec::new(),
                items: text.chars().map(|ch| InputItem::Character(ch)).collect(),
            }
        ];

        for ref input in inputs {
            let input_keys = input.create_inputs()?;
            self.send_keyboard(&input_keys)?;
        }

        Ok(())
        
    }

    /// Simulates starting to hold `keys` on keyboard. Only holdkeys are allowed.
    /// 
    /// The `keys` will be released when `end_hold_keys()` is invoked.
    pub fn begin_hold_keys(&mut self, keys: &str) -> Result<()> {
        let mut holdkeys: Vec<VIRTUAL_KEY> = Vec::new();

        let inputs = Parser::new(false).parse_input(keys)?;
        for input in inputs {
            if input.has_items() {
                return Err(Error::new(ERR_FORMAT, "Error holdkeys"));
            }

            holdkeys.extend(input.holdkeys);
        }

        if holdkeys.is_empty() {
            return Err(Error::new(ERR_FORMAT, "Error holdkeys"));
        }

        let mut holdkey_inputs: Vec<INPUT> = Vec::new();
        for holdkey in &holdkeys {
            holdkey_inputs.push(Input::create_virtual_key(*holdkey, KEYEVENTF_KEYDOWN));
        }
        // send_input(&holdkey_inputs.as_slice())?;
        self.send_keyboard(&holdkey_inputs)?;

        self.holdkeys.extend(holdkeys);

        Ok(())
    }
    
    /// Stop holding keys on keyboard. 
    pub fn end_hold_keys(&mut self) -> Result<()> {
        if self.holdkeys.is_empty() {
            Ok(())
        } else {
            let mut holdkey_inputs = Vec::new();
            for holdkey in self.holdkeys.iter().rev() {
                holdkey_inputs.push(Input::create_virtual_key(*holdkey, KEYEVENTF_KEYUP));
            }
            self.holdkeys.clear();

            // send_input(&holdkey_inputs.as_slice())
            self.send_keyboard(&holdkey_inputs)
        }
    }

    fn send_keyboard(&self, input_keys: &[INPUT]) -> Result<()> {
        // let input_keys = input.create_inputs()?;
        if self.interval == 0 {
            send_input(input_keys)
        } else {
            for input_key in input_keys {
                let input_key_slice: [INPUT; 1] = [input_key.clone()];
                send_input(&input_key_slice)?;

                self.wait();
            }

            Ok(())
        }
    }


    fn wait(&self) {
        if self.interval > 0 {
            sleep(Duration::from_millis(self.interval));
        }
    }
}

impl Drop for Keyboard {
    fn drop(&mut self) {
        if !self.holdkeys.is_empty() {
            let mut holdkey_inputs: Vec<INPUT> = Vec::new();
            for holdkey in self.holdkeys.iter().rev() {
                holdkey_inputs.push(Input::create_virtual_key(*holdkey, KEYEVENTF_KEYUP));
            }

            if send_input(&holdkey_inputs.as_slice()).is_ok() {
                self.holdkeys.clear();
            }
        }
    }
}

/// Mouse buttons.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MouseButton {
    /// left button
    LEFT,
    /// right button
    RIGHT,
    /// middle button
    MIDDLE,
    /// X button
    X
}

impl MouseButton {
    fn down_flag(&self) -> MOUSE_EVENT_FLAGS {
        match self {
            MouseButton::LEFT => MOUSEEVENTF_LEFTDOWN,
            MouseButton::RIGHT => MOUSEEVENTF_RIGHTDOWN,
            MouseButton::MIDDLE => MOUSEEVENTF_MIDDLEDOWN,
            MouseButton::X => MOUSEEVENTF_XDOWN,
        }
    }

    fn up_flag(&self) -> MOUSE_EVENT_FLAGS {
        match self {
            MouseButton::LEFT => MOUSEEVENTF_LEFTUP,
            MouseButton::RIGHT => MOUSEEVENTF_RIGHTUP,
            MouseButton::MIDDLE => MOUSEEVENTF_MIDDLEUP,
            MouseButton::X => MOUSEEVENTF_XUP,
        }
    }
}

/// Simulate mouse event.
#[derive(Debug)]
pub struct Mouse {
    interval: u64,
    move_time: u64,
    auto_move: bool,
    holdkeys: Vec<VIRTUAL_KEY>
}

impl Default for Mouse {
    fn default() -> Self {
        Self { 
            interval: 100, 
            move_time: 500,
            auto_move: true,
            holdkeys: Vec::new()
        }
    }
}

impl Mouse {
    /// Creates a `Mouse` to simulate mouse event.
    pub fn new() -> Self {
        Self::default()
    }

    /// Sets the interval time between events.
    /// 
    /// `interval` is the time number of milliseconds, `100` is default value.
    pub fn interval(mut self, interval: u64) -> Self {
        self.interval = interval;
        self
    }

    /// Sets the mouse move time in millionseconds. `500` is default value.
    pub fn move_time(mut self, move_time: u64) -> Self {
        self.move_time = move_time;
        self
    }

    /// Sets whether move the cursor to the click point automatically. Default is `true`.
    pub fn auto_move(mut self, auto_move: bool) -> Self {
        self.auto_move = auto_move;
        self
    }

    /// Sets the holdkeys when mouse clicks.
    /// 
    /// The holdkeys is quoted by `{}`. For example: `{Shift}`, `{Ctrl}{Alt}`.
    pub fn holdkeys(mut self, holdkeys: &str) -> Self {
        self.holdkeys.clear();

        let mut expr = holdkeys.chars();
        while let Some((items, is_holdkey)) = Parser::new(false).next_input(&mut expr, true).unwrap() {
            if is_holdkey {
                for item in items {
                    if let InputItem::HoldKey(key) = item {
                        self.holdkeys.push(key);
                    }
                }
            }
        }
            
        self
    }

    /// Retrieves the position of the mouse cursor, in screen coordinates.
    pub fn get_cursor_pos() -> Result<Point> {
        let mut pos: Point = Point::default();
        unsafe { GetCursorPos(pos.as_mut())? };
        Ok(pos)
    }

    /// Moves the cursor to the specified screen coordinates. 
    pub fn set_cursor_pos(pos: &Point) -> Result<()> {
        unsafe { SetCursorPos(pos.get_x(), pos.get_y())? };
        Ok(())
    }

    /// Moves the cursor from current position to the `target` position.
    /// 
    /// # Examples
    /// 
    /// ```
    /// use uiautomation::inputs::Mouse;
    /// use uiautomation::types::Point;
    /// 
    /// let mouse = Mouse::new().move_time(800);
    /// mouse.move_to(&Point::new(10, 20)).unwrap();
    /// mouse.move_to(&Point::new(1000,400)).unwrap();
    /// ```
    pub fn move_to(&self, target: &Point) -> Result<()> {
        if self.move_time > 0 {
            let source = Self::get_cursor_pos()?;
            let delta_x = target.get_x() - source.get_x();
            let delta_y = target.get_y() - source.get_y();

            let delta = max(delta_x.abs(), delta_y.abs());
            let steps = delta / 20;
            if steps > 1 {
                let step_x = delta_x / steps;
                let step_y = delta_y / steps;
                let interval = Duration::from_millis(self.move_time / steps as u64);
                for i in 1..steps {
                    let pos = Point::new(
                        source.get_x() + step_x * i, 
                        source.get_y() + step_y * i
                    );
                    Self::mouse_move_event(&pos)?;
                    sleep(interval);
                }
            }
        }

        Self::mouse_move_event(&target)
    }

    /// Simulates a mouse click event with the specified button.
    /// This will click at the current cursor position.
    /// 
    /// # Examples
    /// 
    /// ```no_run
    /// use uiautomation::inputs::{Mouse, MouseButton};
    /// use uiautomation::types::Point;
    /// 
    /// let mouse = Mouse::new();
    /// mouse.move_to(&Point::new(200, 300)).unwrap();
    /// mouse.click_button(MouseButton::LEFT).unwrap();
    /// ```
    pub fn click_button(&self, button: MouseButton) -> Result<()> {
        let mut guard = MouseGuard::new(self);

        self.before_click(&mut guard)?;
        self.mouse_down(button, &mut guard)?;
        self.mouse_up(button, &mut guard)?;
        self.after_click(&mut guard)?;

        Ok(())
    }

    /// Simulates a mouse double click event with the specified button.
    /// This will double click at the current cursor position.
    /// 
    /// # Examples
    /// ```no_run
    /// use uiautomation::inputs::{Mouse, MouseButton};
    /// use uiautomation::types::Point;
    /// 
    /// let mouse = Mouse::new();
    /// mouse.move_to(&Point::new(200, 300)).unwrap();
    /// mouse.double_click_button(MouseButton::LEFT).unwrap();
    /// ```
    pub fn double_click_button(&self, button: MouseButton) -> Result<()> {
        let mut guard = MouseGuard::new(self);

        self.before_click(&mut guard)?;

        self.mouse_down(button, &mut guard)?;
        self.mouse_up(button, &mut guard)?;

        sleep(Duration::from_millis(max(200, self.interval)));

        self.mouse_down(button, &mut guard)?;
        self.mouse_up(button, &mut guard)?;

        self.after_click(&mut guard)?;

        Ok(())
    }

    /// Simulates a mouse click event.
    /// 
    /// # Examples
    /// ```no_run
    /// use uiautomation::inputs::Mouse;
    /// 
    /// let mouse = Mouse::new();
    /// let pos = Mouse::get_cursor_pos().unwrap();
    /// mouse.click(&pos).unwrap();
    /// ```
    pub fn click(&self, pos: &Point) -> Result<()> {
        // self.before_click()?;
        // self.mouse_click_event(MOUSEEVENTF_LEFTDOWN)?;
        // self.mouse_click_event(MOUSEEVENTF_LEFTUP)?;
        // self.after_click()?;

        // Ok(())
        self.move_to(pos)?;
        self.click_button(MouseButton::LEFT)
    }
    
    /// Simulates a mouse double click event.
    /// 
    /// # Examples
    /// ```no_run
    /// use uiautomation::inputs::Mouse;
    /// 
    /// let mouse = Mouse::new();
    /// let pos = Mouse::get_cursor_pos().unwrap();
    /// mouse.double_click(&pos).unwrap();
    /// ```
    pub fn double_click(&self, pos: &Point) -> Result<()> {
        // if self.auto_move {
        //     self.move_to(pos)?;
        // } else {
        //     Mouse::set_cursor_pos(pos)?;
        // }

        // self.before_click()?;

        // self.mouse_click_event(MOUSEEVENTF_LEFTDOWN)?;
        // self.mouse_click_event(MOUSEEVENTF_LEFTUP)?;

        // sleep(Duration::from_millis(max(200, self.interval)));

        // self.mouse_click_event(MOUSEEVENTF_LEFTDOWN)?;
        // self.mouse_click_event(MOUSEEVENTF_LEFTUP)?;

        // self.after_click()?;

        // Ok(())
        self.move_to(pos)?;
        self.double_click_button(MouseButton::LEFT)
    }

    /// Simulates a right mouse click event.
    /// 
    /// # Examples
    /// ```no_run
    /// use uiautomation::inputs::Mouse;
    /// 
    /// let mouse = Mouse::new();
    /// let pos = Mouse::get_cursor_pos().unwrap();
    /// mouse.right_click(&pos).unwrap();
    /// ```
    pub fn right_click(&self, pos: &Point) -> Result<()> {
        // if self.auto_move {
        //     self.move_to(pos)?;
        // } else {
        //     Mouse::set_cursor_pos(pos)?;
        // }

        // self.before_click()?;
        // self.mouse_click_event(MOUSEEVENTF_RIGHTDOWN)?;
        // self.mouse_click_event(MOUSEEVENTF_RIGHTUP)?;
        // self.after_click()?;

        // Ok(())
        self.move_to(pos)?;
        self.click_button(MouseButton::RIGHT)
    }

    /// Simulates dragging the mouse to the specified position with the specified button.
    /// This will drag from the current cursor position to the target position.
    pub fn drag_to(&self, button: MouseButton, pos: &Point) -> Result<()> {
        let mut guard = MouseGuard::new(self);

        self.before_click(&mut guard)?;
        self.mouse_down(button, &mut guard)?;
        self.move_to(pos)?;
        self.mouse_up(button, &mut guard)?;
        self.after_click(&mut guard)?;

        Ok(())
    }

    fn before_click(&self, guard: &mut MouseGuard) -> Result<()> {
        for holdkey in &self.holdkeys {
            let key_input = [Input::create_virtual_key(*holdkey, KEYEVENTF_KEYDOWN)];
            send_input(&key_input)?;
            self.wait();
        }

        guard.holding = true;

        Ok(())
    }

    fn after_click(&self, guard: &mut MouseGuard) -> Result<()> {
        guard.holding = false;

        for holdkey in &self.holdkeys {
            let key_input = [Input::create_virtual_key(*holdkey, KEYEVENTF_KEYUP)];
            send_input(&key_input)?;
            self.wait();
        }

        Ok(())
    }

    fn mouse_down(&self, button: MouseButton, guard: &mut MouseGuard) -> Result<()> {
        self.mouse_click_event(button.down_flag())?;
        guard.button = Some(button);
        Ok(())
    }

    fn mouse_up(&self, button: MouseButton, guard: &mut MouseGuard) -> Result<()> {
        if guard.button == Some(button) {
            self.mouse_click_event(button.up_flag())?;
            guard.button = None;
            Ok(())
        } else {
            Err(Error::new(ERR_INVALID_ARG, "Mouse not holding this button"))
        }
    }

    fn mouse_click_event(&self, flags: MOUSE_EVENT_FLAGS) -> Result<()> {
        let input = [INPUT {
            r#type: INPUT_MOUSE,
            Anonymous: INPUT_0 { 
                mi: MOUSEINPUT { 
                    dx: 0,
                    dy: 0,
                    mouseData: 0, 
                    dwFlags: flags, 
                    time: 0, 
                    dwExtraInfo: 0 
                }
            }
        }];
        send_input(&input)?;
        self.wait();

        Ok(())
    }

    /// Moves the cursor to the specified screen coordinates using send_input().
    fn mouse_move_event(pos: &Point) -> Result<()> {
        let (cx_screen, cy_screen) = get_screen_size()?;
        let dx = (pos.get_x() as f32 / cx_screen as f32 * 65535.0) as i32;
        let dy = (pos.get_y() as f32 / cy_screen as f32 * 65535.0) as i32;
        let input = [INPUT {
            r#type: INPUT_MOUSE,
            Anonymous: INPUT_0 {
                mi: MOUSEINPUT {
                    dx,
                    dy,
                    mouseData: 0,
                    dwFlags: MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE,
                    time: 0,
                    dwExtraInfo: 0,
                },
            },
        }];
        send_input(&input)?;
        Ok(())
    }

    fn wait(&self) {
        if self.interval > 0 {
            sleep(Duration::from_millis(self.interval));
        }
    }
}

struct MouseGuard<'a> {
    mouse: &'a Mouse,
    holding: bool,
    button: Option<MouseButton>,
}

impl<'a> MouseGuard<'a> {
    fn new(mouse: &'a Mouse) -> Self {
        Self { 
            mouse, 
            holding: false, 
            button: None 
        }
    }
}

impl Drop for MouseGuard<'_> {
    fn drop(&mut self) {
        if let Some(button) = self.button {
            self.button = None;
            if let Err(e) = self.mouse.mouse_up(button, self) {
                log_error!("Error releasing mouse button: {}", e);
            }
        }

        if self.holding {
            self.holding = false;
            if let Err(e) = self.mouse.after_click(self) {
                log_error!("Error releasing holdkeys: {}", e);
            }
        }
    }
}

/// Retrieves the `(width, height)` size of the primary screen.
pub fn get_screen_size() -> Result<(i32, i32)> {
    let width = unsafe { GetSystemMetrics(SM_CXSCREEN) };
    if width == 0 {
        return Err(Error::last_os_error());
    }

    let height = unsafe { GetSystemMetrics(SM_CYSCREEN) };
    if height == 0 {
        return Err(Error::last_os_error());
    }

    Ok((width, height))
}

fn send_input(inputs: &[INPUT]) -> Result<()> {
    let sent = unsafe {
        SendInput(inputs, mem::size_of::<INPUT>() as _)
    };

    if sent == inputs.len() as u32 {
        Ok(())
    } else {
        Err(Error::last_os_error())
    }
}

#[cfg(test)]
mod tests {
    use std::vec;

    use windows::Win32::UI::Input::KeyboardAndMouse::*;

    use crate::inputs::get_screen_size;
    use crate::inputs::init_virtual_keys;
    use crate::inputs::Keyboard;
    use crate::inputs::Input;
    use crate::inputs::InputItem;
    use crate::inputs::Parser;
    use crate::inputs::VIRTUAL_KEYS;

    #[test]
    fn test_virtual_keys() {
        let vkeys = VIRTUAL_KEYS.get_or_init(init_virtual_keys);

        let key = vkeys.get("LBUTTON");
        assert_eq!(key, Some(&VK_LBUTTON));
    }

    #[test]
    fn switch_desktop() {
        let kb = Keyboard::default().interval(50);
        kb.send_keys("{win}D").unwrap();

        std::thread::sleep(std::time::Duration::from_millis(500));
        
        kb.send_keys("{ALT}{TAB}").unwrap();
    }

    #[test]
    fn test_parse_input_1() {
        assert_eq!(
            Parser::new(false).parse_input("{ctrl}c").unwrap(),
            vec![Input {
                holdkeys: vec![VK_CONTROL],
                items: vec![InputItem::Character('c')]
            }]
        );
    }

    #[test]
    fn test_parse_input_2() {
        assert_eq!(
            Parser::new(false).parse_input("{ctrl}{alt}{delete}").unwrap(),
            vec![Input {
                holdkeys: vec![VK_CONTROL, VK_MENU],
                items: vec![InputItem::VirtualKey(VK_DELETE, 1)]
            }]
        );
    }

    #[test]
    fn test_parse_input_3() {
        assert_eq!(
            Parser::new(false).parse_input("{shift}(ab)").unwrap(),
            vec![Input {
                holdkeys: vec![VK_SHIFT],
                items: vec![InputItem::Character('a'), InputItem::Character('b')]
            }]
        );
    }

    #[test]
    fn test_parse_input_4() {
        assert_eq!(
            Parser::new(false).parse_input("{{}{}}{(}{)}").unwrap(),
            vec![
                Input {
                    holdkeys: Vec::new(),
                    items: vec![InputItem::Character('{'), InputItem::Character('}'), InputItem::Character('('), InputItem::Character(')')]
                }
            ]
        )
    }

    #[test]
    fn test_parse_input_5() {
        assert!(Parser::new(false).parse_input("Hello,Rust UIAutomation!{enter}").is_ok());
    }

    #[test]
    fn test_parse_input_6() {
        assert_eq!(
            Parser::new(false).parse_input("你好！").unwrap(),
            vec![
                Input {
                    holdkeys: Vec::new(),
                    items: vec![InputItem::Character('你'), InputItem::Character('好'), InputItem::Character('！')]
                }
            ]
        )
    }

    #[test]
    fn test_parse_input_7() {
        assert_eq!(
            Parser::new(true).parse_input("{").unwrap(),
            vec![
                Input {
                    holdkeys: Vec::new(),
                    items: vec![InputItem::Character('{')]
                }
            ]
        );
        assert!(
            Parser::new(false).parse_input("{").is_err()
        );
        assert_eq!(
            Parser::new(true).parse_input("{ctrl}{").unwrap(),
            vec![
                Input {
                    holdkeys: vec![VK_CONTROL],
                    items: vec![InputItem::Character('{')]
                }
            ]
        );
        assert_eq!(
            Parser::new(true).parse_input("{}").unwrap(),
            vec![
                Input {
                    holdkeys: Vec::new(),
                    items: vec![InputItem::Character('{'), InputItem::Character('}')]
                }
            ]
        );
    }

    #[test]
    fn test_parse_input_8() {
        assert_eq!(
            Parser::new(true).parse_input("(").unwrap(),
            vec![
                Input {
                    holdkeys: Vec::new(),
                    items: vec![InputItem::Character('(')]
                }
            ]
        );
        assert!(
            Parser::new(false).parse_input("(").is_ok()
        );
        assert_eq!(
            Parser::new(true).parse_input("({ctrl}()").unwrap(),
            vec![
                Input {
                    holdkeys: Vec::new(),
                    items: vec![InputItem::Character('(')]
                },
                Input {
                    holdkeys: vec![VK_CONTROL],
                    items: vec![InputItem::Character('(')]
                },
                Input {
                    holdkeys: Vec::new(),
                    items: vec![InputItem::Character(')')]
                },
            ]
        );
        assert_eq!(
            Parser::new(true).parse_input("({ctrl}{)").unwrap(),
            vec![
                Input {
                    holdkeys: Vec::new(),
                    items: vec![InputItem::Character('(')]
                },
                Input {
                    holdkeys: vec![VK_CONTROL],
                    items: vec![InputItem::Character('{')]
                },
                Input {
                    holdkeys: Vec::new(),
                    items: vec![InputItem::Character(')')]
                },
            ]
        );
        assert_eq!(
            Parser::new(true).parse_input("()").unwrap(),
            vec![
                Input {
                    holdkeys: Vec::new(),
                    items: vec![InputItem::Character('('), InputItem::Character(')')]
                }
            ]
        );
    }

    #[test]
    fn test_parse_input_9() {
        assert!(
            Parser::new(true).parse_input(" {None} (Keys).").is_ok()
        );
    }

    #[test]
    fn test_repeat_keys() {
        let input = Parser::new(false).parse_input("{ENTER 3}").unwrap();
        assert_eq!(
            input,
            vec![Input {
                holdkeys: vec![],
                items: vec![InputItem::VirtualKey(VK_RETURN, 3)]
            }]
        );
    }

    #[test]
    fn test_zh_input() {
        let inputs = Parser::new(false).parse_input("你好").unwrap();
        for input in &inputs {
            let keys = input.create_inputs();
            assert!(keys.is_ok());
        }
    }

    #[test]
    fn test_screen_size() {
        let (w, h) = get_screen_size().unwrap();
        assert!(w > 0 && h > 0);
        println!("Screen size: {}x{}", w, h);
    }
}
