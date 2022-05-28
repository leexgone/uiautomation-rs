use std::fmt::Debug;
use std::fmt::Display;

use windows::Win32::Foundation::POINT;
use windows::Win32::Foundation::RECT;

/// A Point type stores the x and y position.
#[derive(Clone, Copy, PartialEq, Eq, Default)]
pub struct Point(POINT);

impl Point {
    /// Creates a new position.
    pub fn new(x: i32, y: i32) -> Self {
        Self(POINT {
            x,
            y
        })
    }

    /// Retrievies the x position.
    pub fn get_x(&self) -> i32 {
        self.0.x
    }

    /// Sets the x position.
    pub fn set_x(&mut self, x: i32) {
        self.0.x = x;
    }

    /// Retrieves the y position.
    pub fn get_y(&self) -> i32 {
        self.0.y
    }

    /// Sets the y position.
    pub fn set_y(&mut self, y: i32) {
        self.0.y = y;
    }
}

impl Debug for Point {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Point").field("x", &self.0.x).field("y", &self.0.y).finish()
    }
}

impl Display for Point {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "({}, {})", self.0.x, self.0.y)
    }
}

impl From<POINT> for Point {
    fn from(point: POINT) -> Self {
        Self(point)
    }
}

impl Into<POINT> for Point {
    fn into(self) -> POINT {
        self.0
    }
}

impl AsRef<POINT> for Point {
    fn as_ref(&self) -> &POINT {
        &self.0
    }
}

impl AsMut<POINT> for Point {
    fn as_mut(&mut self) -> &mut POINT {
        &mut self.0
    }
}

/// A Rect type stores the position and size of a rectangle.
#[derive(Clone, Copy, PartialEq, Eq, Default)]
pub struct Rect(RECT);

impl Rect {
    /// Creates a new rect.
    pub fn new(left: i32, top: i32, right: i32, bottom: i32) -> Self {
        Self(RECT {
            left,
            top,
            right,
            bottom
        })
    }

    /// Retrieves the left of the rect.
    pub fn get_left(&self) -> i32 {
        self.0.left
    }

    /// Sets the left of the rect.
    pub fn set_left(&mut self, left: i32) {
        self.0.left = left;
    }

    /// Retrieves the top of the rect.
    pub fn get_top(&self) -> i32 {
        self.0.top
    }

    /// Sets the top of the rect.
    pub fn set_top(&mut self, top: i32) {
        self.0.top = top;
    }

    /// Retrieves the right of the rect.
    pub fn get_right(&self) -> i32 {
        self.0.right
    }

    /// Sets the right of the rect.
    pub fn set_right(&mut self, right: i32) {
        self.0.right = right;
    }

    /// Retrieves the bottom of the rect.
    pub fn get_bottom(&self) -> i32 {
        self.0.bottom
    }

    /// Sets the bottom of the rect.
    pub fn set_bottom(&mut self, bottom: i32) {
        self.0.bottom = bottom;
    }

    /// Retrieves the top left point.
    pub fn get_top_left(&self) -> Point {
        Point::new(self.get_left(), self.get_top())
    }

    /// Retrieves the right bottom point.
    pub fn get_right_bottom(&self) -> Point {
        Point::new(self.get_right(), self.get_bottom())
    }
}

impl Debug for Rect {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Rect").field("left", &self.0.left).field("top", &self.0.top).field("right", &self.0.right).field("bottom", &self.0.bottom).finish()
    }
}

impl Display for Rect {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "[({}, {}), ({}, {})]", self.0.left, self.0.top, self.0.right, self.0.bottom)
    }
}

impl From<RECT> for Rect {
    fn from(rect: RECT) -> Self {
        Self(rect)
    }
}

impl Into<RECT> for Rect {
    fn into(self) -> RECT {
        self.0
    }
}

impl AsRef<RECT> for Rect {
    fn as_ref(&self) -> &RECT {
        &self.0
    }
}

impl AsMut<RECT> for Rect {
    fn as_mut(&mut self) -> &mut RECT {
        &mut self.0
    }
}