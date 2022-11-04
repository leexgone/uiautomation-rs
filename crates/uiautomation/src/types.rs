use std::fmt::Debug;
use std::fmt::Display;

use windows::Win32::Foundation::HWND;
use windows::Win32::Foundation::POINT;
use windows::Win32::Foundation::RECT;
use windows::Win32::UI::Accessibility::UIA_PROPERTY_ID;

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

// impl<'a> IntoParam<'a, POINT> for Point {
//     fn into_param(self) -> windows::core::Param<'a, POINT> {
//         Param::Owned(self.0)
//     }
// }

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

    /// Retrieves the width of the rect.
    pub fn get_width(&self) -> i32 {
        self.0.right - self.0.left + 1
    }

    /// Retrieves the height of the rect.
    pub fn get_height(&self) -> i32 {
        self.0.bottom - self.0.top + 1
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

// impl<'a> IntoParam<'a, RECT> for Rect {
//     fn into_param(self) -> Param<'a, RECT> {
//         Param::Owned(self.0)
//     }
// }

/// A Wrapper for windows `HWND`.
#[derive(Default, Clone, Copy)]
pub struct Handle(HWND);

impl Debug for Handle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Handle(0x{:X})", self.0.0)
    }
}

impl Display for Handle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "0x{:X}", self.0.0)
    }
}

impl From<HWND> for Handle {
    fn from(hwnd: HWND) -> Self {
        Self(hwnd)
    }
}

impl Into<HWND> for Handle {
    fn into(self) -> HWND {
        self.0
    }
}

impl AsRef<HWND> for Handle {
    fn as_ref(&self) -> &HWND {
        &self.0
    }
}

/// Defines enum for `UIA_PROPERTY_ID`
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum UIProperty {
    RuntimeId = 30000u32,
    BoundingRectangle = 30001u32,
    ProcessId = 30002u32,
    ControlType = 30003u32,
    LocalizedControlType = 30004u32,
    Name = 30005u32,
    AcceleratorKey = 30006u32,
    AccessKey = 30007u32,
    HasKeyboardFocus = 30008u32,
    IsKeyboardFocusable = 30009u32,
    IsEnabled = 30010u32,
    AutomationId = 30011u32,
    ClassName = 30012u32,
    HelpText = 30013u32,
    ClickablePoint = 30014u32,
    Culture = 30015u32,
    IsControlElement = 30016u32,
    IsContentElement = 30017u32,
    LabeledBy = 30018u32,
    IsPassword = 30019u32,
    NativeWindowHandle = 30020u32,
    ItemType = 30021u32,
    IsOffscreen = 30022u32,
    Orientation = 30023u32,
    FrameworkId = 30024u32,
    IsRequiredForForm = 30025u32,
    ItemStatus = 30026u32,
    IsDockPatternAvailable = 30027u32,
    IsExpandCollapsePatternAvailable = 30028u32,
    IsGridItemPatternAvailable = 30029u32,
    IsGridPatternAvailable = 30030u32,
    IsInvokePatternAvailable = 30031u32,
    IsMultipleViewPatternAvailable = 30032u32,
    IsRangeValuePatternAvailable = 30033u32,
    IsScrollPatternAvailable = 30034u32,
    IsScrollItemPatternAvailable = 30035u32,
    IsSelectionItemPatternAvailable = 30036u32,
    IsSelectionPatternAvailable = 30037u32,
    IsTablePatternAvailable = 30038u32,
    IsTableItemPatternAvailable = 30039u32,
    IsTextPatternAvailable = 30040u32,
    IsTogglePatternAvailable = 30041u32,
    IsTransformPatternAvailable = 30042u32,
    IsValuePatternAvailable = 30043u32,
    IsWindowPatternAvailable = 30044u32,
    ValueValue = 30045u32,
    ValueIsReadOnly = 30046u32,
    RangeValueValue = 30047u32,
    RangeValueIsReadOnly = 30048u32,
    RangeValueMinimum = 30049u32,
    RangeValueMaximum = 30050u32,
    RangeValueLargeChange = 30051u32,
    RangeValueSmallChange = 30052u32,
    ScrollHorizontalScrollPercent = 30053u32,
    ScrollHorizontalViewSize = 30054u32,
    ScrollVerticalScrollPercent = 30055u32,
    ScrollVerticalViewSize = 30056u32,
    ScrollHorizontallyScrollable = 30057u32,
    ScrollVerticallyScrollable = 30058u32,
    SelectionSelection = 30059u32,
    SelectionCanSelectMultiple = 30060u32,
    SelectionIsSelectionRequired = 30061u32,
    GridRowCount = 30062u32,
    GridColumnCount = 30063u32,
    GridItemRow = 30064u32,
    GridItemColumn = 30065u32,
    GridItemRowSpan = 30066u32,
    GridItemColumnSpan = 30067u32,
    GridItemContainingGrid = 30068u32,
    DockDockPosition = 30069u32,
    ExpandCollapseExpandCollapseState = 30070u32,
    MultipleViewCurrentView = 30071u32,
    MultipleViewSupportedViews = 30072u32,
    WindowCanMaximize = 30073u32,
    WindowCanMinimize = 30074u32,
    WindowWindowVisualState = 30075u32,
    WindowWindowInteractionState = 30076u32,
    WindowIsModal = 30077u32,
    WindowIsTopmost = 30078u32,
    SelectionItemIsSelected = 30079u32,
    SelectionItemSelectionContainer = 30080u32,
    TableRowHeaders = 30081u32,
    TableColumnHeaders = 30082u32,
    TableRowOrColumnMajor = 30083u32,
    TableItemRowHeaderItems = 30084u32,
    TableItemColumnHeaderItems = 30085u32,
    ToggleToggleState = 30086u32,
    TransformCanMove = 30087u32,
    TransformCanResize = 30088u32,
    TransformCanRotate = 30089u32,
    IsLegacyIAccessiblePatternAvailable = 30090u32,
    LegacyIAccessibleChildId = 30091u32,
    LegacyIAccessibleName = 30092u32,
    LegacyIAccessibleValue = 30093u32,
    LegacyIAccessibleDescription = 30094u32,
    LegacyIAccessibleRole = 30095u32,
    LegacyIAccessibleState = 30096u32,
    LegacyIAccessibleHelp = 30097u32,
    LegacyIAccessibleKeyboardShortcut = 30098u32,
    LegacyIAccessibleSelection = 30099u32,
    LegacyIAccessibleDefaultAction = 30100u32,
    AriaRole = 30101u32,
    AriaProperties = 30102u32,
    IsDataValidForForm = 30103u32,
    ControllerFor = 30104u32,
    DescribedBy = 30105u32,
    FlowsTo = 30106u32,
    ProviderDescription = 30107u32,
    IsItemContainerPatternAvailable = 30108u32,
    IsVirtualizedItemPatternAvailable = 30109u32,
    IsSynchronizedInputPatternAvailable = 30110u32,
    OptimizeForVisualContent = 30111u32,
    IsObjectModelPatternAvailable = 30112u32,
    AnnotationAnnotationTypeId = 30113u32,
    AnnotationAnnotationTypeName = 30114u32,
    AnnotationAuthor = 30115u32,
    AnnotationDateTime = 30116u32,
    AnnotationTarget = 30117u32,
    IsAnnotationPatternAvailable = 30118u32,
    IsTextPattern2Available = 30119u32,
    StylesStyleId = 30120u32,
    StylesStyleName = 30121u32,
    StylesFillColor = 30122u32,
    StylesFillPatternStyle = 30123u32,
    StylesShape = 30124u32,
    StylesFillPatternColor = 30125u32,
    StylesExtendedProperties = 30126u32,
    IsStylesPatternAvailable = 30127u32,
    IsSpreadsheetPatternAvailable = 30128u32,
    SpreadsheetItemFormula = 30129u32,
    SpreadsheetItemAnnotationObjects = 30130u32,
    SpreadsheetItemAnnotationTypes = 30131u32,
    IsSpreadsheetItemPatternAvailable = 30132u32,
    Transform2CanZoom = 30133u32,
    IsTransformPattern2Available = 30134u32,
    LiveSetting = 30135u32,
    IsTextChildPatternAvailable = 30136u32,
    IsDragPatternAvailable = 30137u32,
    DragIsGrabbed = 30138u32,
    DragDropEffect = 30139u32,
    DragDropEffects = 30140u32,
    IsDropTargetPatternAvailable = 30141u32,
    DropTargetDropTargetEffect = 30142u32,
    DropTargetDropTargetEffects = 30143u32,
    DragGrabbedItems = 30144u32,
    Transform2ZoomLevel = 30145u32,
    Transform2ZoomMinimum = 30146u32,
    Transform2ZoomMaximum = 30147u32,
    FlowsFrom = 30148u32,
    IsTextEditPatternAvailable = 30149u32,
    IsPeripheral = 30150u32,
    IsCustomNavigationPatternAvailable = 30151u32,
    PositionInSet = 30152u32,
    SizeOfSet = 30153u32,
    Level = 30154u32,
    AnnotationTypes = 30155u32,
    AnnotationObjects = 30156u32,
    LandmarkType = 30157u32,
    LocalizedLandmarkType = 30158u32,
    FullDescription = 30159u32,
    FillColor = 30160u32,
    OutlineColor = 30161u32,
    FillType = 30162u32,
    VisualEffects = 30163u32,
    OutlineThickness = 30164u32,
    CenterPoint = 30165u32,
    Rotation = 30166u32,
    Size = 30167u32,
    IsSelectionPattern2Available = 30168u32,
    Selection2FirstSelectedItem = 30169u32,
    Selection2LastSelectedItem = 30170u32,
    Selection2CurrentSelectedItem = 30171u32,
    Selection2ItemCount = 30172u32,
    HeadingLevel = 30173u32,
    IsDialog = 30174u32,
}

// impl From<UIA_PROPERTY_ID> for UIProperty {
//     fn from(prop_id: UIA_PROPERTY_ID) -> Self {
//         prop_id.0 as _
//     }
// }