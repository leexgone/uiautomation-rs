use std::ffi::c_void;
use std::fmt::Debug;
use std::fmt::Display;

use uiautomation_derive::EnumConvert;
use uiautomation_derive::map_as;
use windows::core::Param;
use windows::Win32::Foundation::HWND;
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

/// A Wrapper for windows `HWND`.
#[derive(Default, Clone, Copy)]
pub struct Handle(HWND);

impl Debug for Handle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let v: i64 = unsafe {
            std::mem::transmute(self.0.0)
        };
        write!(f, "Handle(0x{:X})", v)
    }
}

impl Display for Handle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{:?}", self.0)
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

impl Param<HWND> for Handle {
    unsafe fn param(self) -> windows::core::ParamValue<HWND> {
        windows::core::ParamValue::Owned(self.0)
    }
}

impl From<i64> for Handle {
    fn from(value: i64) -> Self {
        let hwd: *mut c_void = unsafe {
            std::mem::transmute(value)
        };
        Self(HWND(hwd))
    }
}

impl Into<i64> for Handle {
    fn into(self) -> i64 {
        unsafe {
            std::mem::transmute(self.0.0)
        }
    }
}

// impl AsRef<isize> for Handle {
//     fn as_ref(&self) -> &isize {
//         &self.0.0
//     }
// }

/// Defines enum for `windows::Win32::UI::Accessibility::UIA_PROPERTY_ID`.
/// 
/// Describes the named constants that identify the properties of Microsoft UI Automation elements.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::UIA_PROPERTY_ID)]
pub enum UIProperty {
    /// Identifies the RuntimeId property, which is an array of integers representing the identifier for an automation element.
    RuntimeId = 30000i32,
    /// Identifies the BoundingRectangle property, which specifies the coordinates of the rectangle that completely encloses the automation element.
    BoundingRectangle = 30001i32,
    /// Identifies the ProcessId property, which is an integer representing the process identifier (ID) of the automation element.
    ProcessId = 30002i32,
    /// Identifies the ControlType property, which is a class that identifies the type of the automation element. 
    ControlType = 30003i32,
    /// Identifies the LocalizedControlType property, which is a text string describing the type of control that the automation element represents.
    LocalizedControlType = 30004i32,
    /// Identifies the Name property, which is a string that holds the name of the automation element.
    Name = 30005i32,
    /// Identifies the AcceleratorKey property, which is a string containing the accelerator key (also called shortcut key) combinations for the automation element.
    AcceleratorKey = 30006i32,
    /// Identifies the AccessKey property, which is a string containing the access key character for the automation element.
    AccessKey = 30007i32,
    /// Identifies the HasKeyboardFocus property, which is a Boolean value that indicates whether the automation element has keyboard focus.
    HasKeyboardFocus = 30008i32,
    /// Identifies the IsKeyboardFocusable property, which is a Boolean value that indicates whether the automation element can accept keyboard focus.
    IsKeyboardFocusable = 30009i32,
    /// Identifies the IsEnabled property, which is a Boolean value that indicates whether the UI item referenced by the automation element is enabled and can be interacted with.
    IsEnabled = 30010i32,
    /// Identifies the AutomationId property, which is a string containing the UI Automation identifier (ID) for the automation element.
    AutomationId = 30011i32,
    /// Identifies the ClassName property, which is a string containing the class name for the automation element as assigned by the control developer.
    ClassName = 30012i32,
    /// Identifies the HelpText property, which is a help text string associated with the automation element.
    HelpText = 30013i32,
    /// Identifies the ClickablePoint property, which is a point on the automation element that can be clicked. 
    ClickablePoint = 30014i32,
    /// Identifies the Culture property, which contains a locale identifier for the automation element.
    Culture = 30015i32,
    /// Identifies the IsControlElement property, which is a Boolean value that specifies whether the element appears in the control view of the automation element tree.
    IsControlElement = 30016i32,
    /// Identifies the IsContentElement property, which is a Boolean value that specifies whether the element appears in the content view of the automation element tree.
    IsContentElement = 30017i32,
    /// Identifies the LabeledBy property, which is an automation element that contains the text label for this element.
    LabeledBy = 30018i32,
    /// Identifies the IsPassword property, which is a Boolean value that indicates whether the automation element contains protected content or a password.
    IsPassword = 30019i32,
    /// Identifies the NativeWindowHandle property, which is an integer that represents the handle (HWND) of the automation element window, if it exists; otherwise, this property is 0.
    NativeWindowHandle = 30020i32,
    /// Identifies the ItemType property, which is a text string describing the type of the automation element.
    ItemType = 30021i32,
    /// Identifies the IsOffscreen property, which is a Boolean value that indicates whether the automation element is entirely scrolled out of view (for example, an item in a list box that is outside the viewport of the container object) or collapsed out of view (for example, an item in a tree view or menu, or in a minimized window). 
    IsOffscreen = 30022i32,
    /// Identifies the Orientation property, which indicates the orientation of the control represented by the automation element. The property is expressed as a value from the OrientationType enumerated type.
    Orientation = 30023i32,
    /// Identifies the FrameworkId property, which is a string containing the name of the underlying UI framework that the automation element belongs to.
    FrameworkId = 30024i32,
    /// Identifies the IsRequiredForForm property, which is a Boolean value that indicates whether the automation element is required to be filled out on a form.
    IsRequiredForForm = 30025i32,
    /// Identifies the ItemStatus property, which is a text string describing the status of an item of the automation element.
    ItemStatus = 30026i32,
    /// Identifies the IsDockPatternAvailable property, which indicates whether the Dock control pattern is available for the automation element.
    IsDockPatternAvailable = 30027i32,
    /// Identifies the IsExpandCollapsePatternAvailable property, which indicates whether the ExpandCollapse control pattern is available for the automation element.
    IsExpandCollapsePatternAvailable = 30028i32,
    /// Identifies the IsGridItemPatternAvailable property, which indicates whether the GridItem control pattern is available for the automation element.
    IsGridItemPatternAvailable = 30029i32,
    /// Identifies the IsGridPatternAvailable property, which indicates whether the Grid control pattern is available for the automation element.
    IsGridPatternAvailable = 30030i32,
    /// Identifies the IsInvokePatternAvailable property, which indicates whether the Invoke control pattern is available for the automation element.
    IsInvokePatternAvailable = 30031i32,
    /// Identifies the IsMultipleViewPatternAvailable property, which indicates whether the MultipleView control pattern is available for the automation element.
    IsMultipleViewPatternAvailable = 30032i32,
    /// Identifies the IsRangeValuePatternAvailable property, which indicates whether the RangeValue control pattern is available for the automation element.
    IsRangeValuePatternAvailable = 30033i32,
    /// Identifies the IsScrollPatternAvailable property, which indicates whether the Scroll control pattern is available for the automation element.
    IsScrollPatternAvailable = 30034i32,
    /// Identifies the IsScrollItemPatternAvailable property, which indicates whether the ScrollItem control pattern is available for the automation element.
    IsScrollItemPatternAvailable = 30035i32,
    /// Identifies the IsSelectionItemPatternAvailable property, which indicates whether the SelectionItem control pattern is available for the automation element. 
    IsSelectionItemPatternAvailable = 30036i32,
    /// Identifies the IsSelectionPatternAvailable property, which indicates whether the Selection control pattern is available for the automation element.
    IsSelectionPatternAvailable = 30037i32,
    /// Identifies the IsTablePatternAvailable property, which indicates whether the Table control pattern is available for the automation element.
    IsTablePatternAvailable = 30038i32,
    /// Identifies the IsTableItemPatternAvailable property, which indicates whether the TableItem control pattern is available for the automation element.
    IsTableItemPatternAvailable = 30039i32,
    /// Identifies the IsTextPatternAvailable property, which indicates whether the Text control pattern is available for the automation element.
    IsTextPatternAvailable = 30040i32,
    /// Identifies the IsTogglePatternAvailable property, which indicates whether the Toggle control pattern is available for the automation element.
    IsTogglePatternAvailable = 30041i32,
    /// Identifies the IsTransformPatternAvailable property, which indicates whether the Transform control pattern is available for the automation element.
    IsTransformPatternAvailable = 30042i32,
    /// Identifies the IsValuePatternAvailable property, which indicates whether the Value control pattern is available for the automation element.
    IsValuePatternAvailable = 30043i32,
    /// Identifies the IsWindowPatternAvailable property, which indicates whether the Window control pattern is available for the automation element.
    IsWindowPatternAvailable = 30044i32,
    /// Identifies the Value property of the Value control pattern.
    ValueValue = 30045i32,
    /// Identifies the IsReadOnly property of the Value control pattern.
    ValueIsReadOnly = 30046i32,
    /// Identifies the Value property of the RangeValue control pattern.
    RangeValueValue = 30047i32,
    /// Identifies the IsReadOnly property of the RangeValue control pattern.
    RangeValueIsReadOnly = 30048i32,
    /// Identifies the Minimum property of the RangeValue control pattern.
    RangeValueMinimum = 30049i32,
    /// Identifies the Maximum property of the RangeValue control pattern.
    RangeValueMaximum = 30050i32,
    /// Identifies the LargeChange property of the RangeValue control pattern.
    RangeValueLargeChange = 30051i32,
    /// Identifies the SmallChange property of the RangeValue control pattern.
    RangeValueSmallChange = 30052i32,
    /// Identifies the HorizontalScrollPercent property of the Scroll control pattern.
    ScrollHorizontalScrollPercent = 30053i32,
    /// Identifies the HorizontalViewSize property of the Scroll control pattern.
    ScrollHorizontalViewSize = 30054i32,
    /// Identifies the VerticalScrollPercent property of the Scroll control pattern.
    ScrollVerticalScrollPercent = 30055i32,
    /// Identifies the VerticalViewSize property of the Scroll control pattern.
    ScrollVerticalViewSize = 30056i32,
    /// Identifies the HorizontallyScrollable property of the Scroll control pattern.
    ScrollHorizontallyScrollable = 30057i32,
    /// Identifies the VerticallyScrollable property of the Scroll control pattern.
    ScrollVerticallyScrollable = 30058i32,
    /// Identifies the Selection property of the Selection control pattern.
    SelectionSelection = 30059i32,
    /// Identifies the CanSelectMultiple property of the Selection control pattern.
    SelectionCanSelectMultiple = 30060i32,
    /// Identifies the IsSelectionRequired property of the Selection control pattern.
    SelectionIsSelectionRequired = 30061i32,
    /// Identifies the RowCount property of the Grid control pattern. This property indicates the total number of rows in the grid.
    GridRowCount = 30062i32,
    /// Identifies the ColumnCount property of the Grid control pattern. This property indicates the total number of columns in the grid.
    GridColumnCount = 30063i32,
    /// Identifies the Row property of the GridItem control pattern. This property is the ordinal number of the row that contains the cell or item.
    GridItemRow = 30064i32,
    /// Identifies the Column property of the GridItem control pattern. This property indicates the ordinal number of the column that contains the cell or item.
    GridItemColumn = 30065i32,
    /// Identifies the RowSpan property of the GridItem control pattern. This property indicates the number of rows spanned by the cell or item.
    GridItemRowSpan = 30066i32,
    /// Identifies the ColumnSpan property of the GridItem control pattern. This property indicates the number of columns spanned by the cell or item.
    GridItemColumnSpan = 30067i32,
    /// Identifies the ContainingGrid property of the GridItem control pattern.
    GridItemContainingGrid = 30068i32,
    /// Identifies the DockPosition property of the Dock control pattern.
    DockDockPosition = 30069i32,
    /// Identifies the ExpandCollapseState property of the ExpandCollapse control pattern.
    ExpandCollapseExpandCollapseState = 30070i32,
    /// Identifies the CurrentView property of the MultipleView control pattern.
    MultipleViewCurrentView = 30071i32,
    /// Identifies the SupportedViews property of the MultipleView control pattern.
    MultipleViewSupportedViews = 30072i32,
    /// Identifies the CanMaximize property of the Window control pattern.
    WindowCanMaximize = 30073i32,
    /// Identifies the CanMinimize property of the Window control pattern.
    WindowCanMinimize = 30074i32,
    /// Identifies the WindowVisualState property of the Window control pattern.
    WindowWindowVisualState = 30075i32,
    /// Identifies the WindowInteractionState property of the Window control pattern.
    WindowWindowInteractionState = 30076i32,
    /// Identifies the IsModal property of the Window control pattern.
    WindowIsModal = 30077i32,
    /// Identifies the IsTopmost property of the Window control pattern.
    WindowIsTopmost = 30078i32,
    /// Identifies the IsSelected property of the SelectionItem control pattern.
    SelectionItemIsSelected = 30079i32,
    /// Identifies the SelectionContainer property of the SelectionItem control pattern.
    SelectionItemSelectionContainer = 30080i32,
    /// Identifies the RowHeaders property of the Table control pattern.
    TableRowHeaders = 30081i32,
    /// Identifies the ColumnHeaders property of the Table control pattern.
    TableColumnHeaders = 30082i32,
    /// Identifies the RowOrColumnMajor property of the Table control pattern.
    TableRowOrColumnMajor = 30083i32,
    /// Identifies the RowHeaderItems property of the TableItem control pattern.
    TableItemRowHeaderItems = 30084i32,
    /// Identifies the ColumnHeaderItems property of the TableItem control pattern.
    TableItemColumnHeaderItems = 30085i32,
    /// Identifies the ToggleState property of the Toggle control pattern.
    ToggleToggleState = 30086i32,
    /// Identifies the CanMove property of the Transform control pattern.
    TransformCanMove = 30087i32,
    /// Identifies the CanResize property of the Transform control pattern.
    TransformCanResize = 30088i32,
    /// Identifies the CanRotate property of the Transform control pattern.
    TransformCanRotate = 30089i32,
    /// Identifies the IsLegacyIAccessiblePatternAvailable property, which indicates whether the LegacyIAccessible control pattern is available for the automation element.
    IsLegacyIAccessiblePatternAvailable = 30090i32,
    /// Identifies the ChildId property of the LegacyIAccessible control pattern.
    LegacyIAccessibleChildId = 30091i32,
    /// Identifies the Name property of the LegacyIAccessible control pattern.
    LegacyIAccessibleName = 30092i32,
    /// Identifies the Value property of the LegacyIAccessible control pattern.
    LegacyIAccessibleValue = 30093i32,
    /// Identifies the Description property of the LegacyIAccessible control pattern.
    LegacyIAccessibleDescription = 30094i32,
    /// Identifies the Roleproperty of the LegacyIAccessible control pattern.
    LegacyIAccessibleRole = 30095i32,
    /// Identifies the State property of the LegacyIAccessible control pattern.
    LegacyIAccessibleState = 30096i32,
    /// Identifies the Help property of the LegacyIAccessible control pattern.
    LegacyIAccessibleHelp = 30097i32,
    /// Identifies the KeyboardShortcut property of the LegacyIAccessible control pattern.
    LegacyIAccessibleKeyboardShortcut = 30098i32,
    /// Identifies the Selection property of the LegacyIAccessible control pattern.
    LegacyIAccessibleSelection = 30099i32,
    /// Identifies the DefaultAction property of the LegacyIAccessible control pattern.
    LegacyIAccessibleDefaultAction = 30100i32,
    /// Identifies the AriaRole property, which is a string containing the Accessible Rich Internet Application (ARIA) role information for the automation element. 
    AriaRole = 30101i32,
    /// Identifies the AriaProperties property, which is a formatted string containing the Accessible Rich Internet Application (ARIA) property information for the automation element. 
    AriaProperties = 30102i32,
    /// Identifies the IsDataValidForForm property, which is a Boolean value that indicates whether the entered or selected value is valid for the form rule associated with the automation element. 
    IsDataValidForForm = 30103i32,
    /// Identifies the ControllerFor property, which is an array of automation elements that are manipulated by the automation element that supports this property.
    ControllerFor = 30104i32,
    /// Identifies the DescribedBy property, which is an array of elements that provide more information about the automation element.
    DescribedBy = 30105i32,
    /// Identifies the FlowsTo property, which is an array of automation elements that suggests the reading order after the current automation element.
    FlowsTo = 30106i32,
    /// Identifies the ProviderDescription property, which is a formatted string containing the source information of the UI Automation provider for the automation element, including proxy information.
    ProviderDescription = 30107i32,
    /// Identifies the IsItemContainerPatternAvailable property, which indicates whether the ItemContainer control pattern is available for the automation element.
    IsItemContainerPatternAvailable = 30108i32,
    /// Identifies the IsVirtualizedItemPatternAvailable property, which indicates whether the VirtualizedItem control pattern is available for the automation element.
    IsVirtualizedItemPatternAvailable = 30109i32,
    /// Identifies the IsSynchronizedInputPatternAvailable property, which indicates whether the SynchronizedInput control pattern is available for the automation element.
    IsSynchronizedInputPatternAvailable = 30110i32,
    /// Identifies the OptimizeForVisualContent property, which is a Boolean value that indicates whether the provider exposes only elements that are visible.
    OptimizeForVisualContent = 30111i32,
    /// Identifies the IsObjectModelPatternAvailable property, which indicates whether the ObjectModel control pattern is available for the automation element.
    IsObjectModelPatternAvailable = 30112i32,
    /// Identifies the AnnotationTypeId property of the Annotation control pattern. Supported starting with Windows 8.
    AnnotationAnnotationTypeId = 30113i32,
    /// Identifies the AnnotationTypeName property of the Annotation control pattern. Supported starting with Windows 8.
    AnnotationAnnotationTypeName = 30114i32,
    /// Identifies the Author property of the Annotation control pattern. Supported starting with Windows 8.
    AnnotationAuthor = 30115i32,
    /// Identifies the DateTime property of the Annotation control pattern. Supported starting with Windows 8.
    AnnotationDateTime = 30116i32,
    /// Identifies the Target property of the Annotation control pattern. Supported starting with Windows 8.
    AnnotationTarget = 30117i32,
    /// Identifies the IsAnnotationPatternAvailable property, which indicates whether the Annotation control pattern is available for the automation element.
    IsAnnotationPatternAvailable = 30118i32,
    /// Identifies the IsTextPattern2Available property, which indicates whether version two of the Text control pattern is available for the automation element.
    IsTextPattern2Available = 30119i32,
    /// Identifies the StyleId property of the Styles control pattern.
    StylesStyleId = 30120i32,
    /// Identifies the StyleName property of the Styles control pattern.
    StylesStyleName = 30121i32,
    /// Identifies the FillColor property of the Styles control pattern.
    StylesFillColor = 30122i32,
    /// Identifies the FillPatternStyle property of the Styles control pattern.
    StylesFillPatternStyle = 30123i32,
    /// Identifies the Shape property of the Styles control pattern.
    StylesShape = 30124i32,
    /// Identifies the FillPatternColor property of the Styles control pattern.
    StylesFillPatternColor = 30125i32,
    /// Identifies the ExtendedProperties property of the Styles control pattern.
    StylesExtendedProperties = 30126i32,
    /// Identifies the IsStylesPatternAvailable property, which indicates whether the Styles control pattern is available for the automation element.
    IsStylesPatternAvailable = 30127i32,
    /// Identifies the IsSpreadsheetPatternAvailable property, which indicates whether the Spreadsheet control pattern is available for the automation element.
    IsSpreadsheetPatternAvailable = 30128i32,
    /// Identifies the Formula property of the SpreadsheetItem control pattern.
    SpreadsheetItemFormula = 30129i32,
    /// Identifies the AnnotationObjects property of the SpreadsheetItem control pattern.
    SpreadsheetItemAnnotationObjects = 30130i32,
    /// Identifies the AnnotationTypes property of the SpreadsheetItem control pattern. Supported starting with Windows 8.
    SpreadsheetItemAnnotationTypes = 30131i32,
    /// Identifies the IsSpreadsheetItemPatternAvailable property, which indicates whether the SpreadsheetItem control pattern is available for the automation element.
    IsSpreadsheetItemPatternAvailable = 30132i32,
    /// Identifies the CanZoom property of the Transform control pattern. Supported starting with Windows 8.
    Transform2CanZoom = 30133i32,
    /// Identifies the IsTransformPattern2Available property, which indicates whether version two of the Transform control pattern is available for the automation element.
    IsTransformPattern2Available = 30134i32,
    /// Identifies the LiveSetting property, which is supported by an automation element that represents a live region.
    LiveSetting = 30135i32,
    /// Identifies the IsTextChildPatternAvailable property, which indicates whether the TextChild control pattern is available for the automation element.
    IsTextChildPatternAvailable = 30136i32,
    /// Identifies the IsDragPatternAvailable property, which indicates whether the Drag control pattern is available for the automation element.
    IsDragPatternAvailable = 30137i32,
    /// Identifies the IsGrabbed property of the Drag control pattern. Supported starting with Windows 8.
    DragIsGrabbed = 30138i32,
    /// Identifies the DropEffect property of the Drag control pattern. Supported starting with Windows 8.
    DragDropEffect = 30139i32,
    /// Identifies the DropEffects property of the Drag control pattern. Supported starting with Windows 8.
    DragDropEffects = 30140i32,
    /// Identifies the IsDropTargetPatternAvailable property, which indicates whether the DropTarget control pattern is available for the automation element.
    IsDropTargetPatternAvailable = 30141i32,
    /// Identifies the DropTargetEffect property of the DropTarget control pattern. Supported starting with Windows 8.
    DropTargetDropTargetEffect = 30142i32,
    /// Identifies the DropTargetEffects property of the DropTarget control pattern. Supported starting with Windows 8.
    DropTargetDropTargetEffects = 30143i32,
    /// Identifies the GrabbedItems property of the Drag control pattern. Supported starting with Windows 8.
    DragGrabbedItems = 30144i32,
    /// Identifies the ZoomLevel property of the Transform control pattern. Supported starting with Windows 8.
    Transform2ZoomLevel = 30145i32,
    /// Identifies the ZoomMinimum property of the Transform control pattern. Supported starting with Windows 8.
    Transform2ZoomMinimum = 30146i32,
    /// Identifies the ZoomMaximum property of the Transform control pattern. Supported starting with Windows 8.
    Transform2ZoomMaximum = 30147i32,
    /// Identifies the FlowsFrom property, which is an array of automation elements that suggests the reading order before the current automation element. Supported starting with Windows 8.
    FlowsFrom = 30148i32,
    /// Identifies the IsTextEditPatternAvailable property, which indicates whether the TextEdit control pattern is available for the automation element.
    IsTextEditPatternAvailable = 30149i32,
    /// Identifies the IsPeripheral property, which is a Boolean value that indicates whether the automation element represents peripheral UI.
    IsPeripheral = 30150i32,
    /// Identifies the IsCustomNavigationPatternAvailable property, which indicates whether the CustomNavigation control pattern is available for the automation element.
    IsCustomNavigationPatternAvailable = 30151i32,
    /// Identifies the PositionInSet property, which is a 1-based integer associated with an automation element.
    PositionInSet = 30152i32,
    /// Identifies the SizeOfSet property, which is a 1-based integer associated with an automation element.
    SizeOfSet = 30153i32,
    /// Identifies the Level property, which is a 1-based integer associated with an automation element.
    Level = 30154i32,
    /// Identifies the AnnotationTypes property, which is a list of the types of annotations in a document, such as comment, header, footer, and so on.
    AnnotationTypes = 30155i32,
    /// Identifies the AnnotationObjects property, which is a list of annotation objects in a document, such as comment, header, footer, and so on.
    AnnotationObjects = 30156i32,
    /// Identifies the LandmarkType property, which is a Landmark Type Identifier associated with an element.
    LandmarkType = 30157i32,
    /// Identifies the LocalizedLandmarkType, which is a text string describing the type of landmark that the automation element represents.
    LocalizedLandmarkType = 30158i32,
    /// The FullDescription property exposes a localized string which can contain extended description text for an element. 
    FullDescription = 30159i32,
    /// Identifies the FillColor property, which specifies the color used to fill the automation element.
    FillColor = 30160i32,
    /// Identifies the OutlineColor property, which specifies the color used for the outline of the automation element.
    OutlineColor = 30161i32,
    /// Identifies the FillType property, which specifies the pattern used to fill the automation element, such as none, color, gradient, picture, pattern, and so on.
    FillType = 30162i32,
    /// Identifies the VisualEffects property, which is a bit field that specifies effects on the automation element, such as shadow, reflection, glow, soft edges, or bevel.
    VisualEffects = 30163i32,
    /// Identifies the OutlineThickness property, which specifies the width for the outline of the automation element.
    OutlineThickness = 30164i32,
    /// Identifies the CenterPoint property, which specifies the center X and Y point coordinates of the automation element. 
    CenterPoint = 30165i32,
    /// Identifies the Rotation property, which specifies the angle of rotation in unspecified units.
    Rotation = 30166i32,
    /// Identifies the Size property, which specifies the width and height of the automation element.
    Size = 30167i32,
    /// Identifies whether the Selection2 control pattern is available.
    IsSelectionPattern2Available = 30168i32,
    /// Identifies the FirstSelectedItem property of the Selection2 control pattern.
    Selection2FirstSelectedItem = 30169i32,
    /// Identifies the LastSelectedItem property of the Selection2 control pattern.
    Selection2LastSelectedItem = 30170i32,
    /// Identifies the CurrentSelectedItem property of the Selection2 control pattern.
    Selection2CurrentSelectedItem = 30171i32,
    /// Identifies the ItemCount property of the Selection2 control pattern.
    Selection2ItemCount = 30172i32,
    /// Identifies the HeadingLevel property, which indicates the heading level of a UI Automation element.
    HeadingLevel = 30173i32,
    /// Identifies the IsDialog property, which is a Boolean value that indicates whether the automation element is a dialog window. 
    IsDialog = 30174i32,
}

/// Defines enum for `windows::Win32::UI::Accessibility::WindowInteractionState`.
/// 
/// Contains values that specify the current state of the window for purposes of user interaction.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::WindowInteractionState)]
pub enum WindowInteractionState {
    /// The window is running. This does not guarantee that the window is ready for user interaction or is responding.
    Running = 0i32, 
    /// The window is closing.
    Closing = 1i32,
    /// The window is ready for user interaction.
    ReadyForUserInteraction = 2i32,
    /// The window is blocked by a modal window.
    BlockedByModalWindow = 3i32,
    /// The window is not responding.
    NotResponding = 4i32
}

// impl Display for WindowInteractionState {
//     fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
//         match *self {
//             Self::Running => write!(f, "Running"),
//             Self::Closing => write!(f, "Closing"),
//             Self::ReadyForUserInteraction => write!(f, "ReadyForUserInteraction"),
//             Self::BlockedByModalWindow => write!(f, "BlockedByModalWindow"),
//             Self::NotResponding => write!(f, "NotResponding")
//         }
//     }
// }

/// Defines enum for `windows::Win32::UI::Accessibility::DockPosition`.
/// 
/// Contains values that specify the dock position of an object, represented by a DockPattern, within a docking container.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::DockPosition)]
pub enum DockPosition {
    /// Indicates that the UI Automation element is docked along the top edge of the docking container.
    Top = 0i32,
    /// Indicates that the UI Automation element is docked along the left edge of the docking container.
    Left = 1i32,
    /// Indicates that the UI Automation element is docked along the bottom edge of the docking container.
    Bottom = 2i32,
    /// Indicates that the UI Automation element is docked along the right edge of the docking container.
    Right = 3i32,
    /// Indicates that the UI Automation element is docked along all edges of the docking container and fills all available space within the container.
    Fill = 4i32,
    /// Indicates that the UI Automation element is not docked to any edge of the docking container.
    None = 5i32,
}

/// Defines enum for `windows::Win32::UI::Accessibility::ExpandCollapseState`.
/// 
/// Contains values that specify the ExpandCollapseState automation property value of a UI Automation element.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::ExpandCollapseState)]
pub enum ExpandCollapseState {
    /// No child nodes, controls, or content of the UI Automation element are displayed.
    Collapsed = 0i32,
    /// All child nodes, controls, and content of the UI Automation element are displayed.
    Expanded = 1i32,
    /// Some, but not all, child nodes, controls, or content of the UI Automation element are displayed.
    PartiallyExpanded = 2i32,
    /// The UI Automation element has no child nodes, controls, or content to display.
    LeafNode = 3i32
}

/// Defines enum for `windows::Win32::UI::Accessibility::NavigateDirection`.
/// 
/// Contains values used to specify the direction of navigation within the Microsoft UI Automation tree.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::NavigateDirection)]
pub enum NavigateDirection {
    /// The navigation direction is to the parent.
    Parent = 0i32,
    /// The navigation direction is to the next sibling.
    NextSibling = 1i32,
    /// The navigation direction is to the previous sibling.
    PreviousSibling = 2i32,
    /// The navigation direction is to the first child.
    FirstChild = 3i32,
    /// The navigation direction is to the last child.
    LastChild = 4i32
}

/// Defines enum for `windows::Win32::UI::Accessibility::RowOrColumnMajor`.
/// 
/// Contains values that specify whether data in a table should be read primarily by row or by column.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::RowOrColumnMajor)]
pub enum RowOrColumnMajor {
    /// Data in the table should be read row by row.
    RowMajor = 0i32,
    /// Data in the table should be read column by column.
    ColumnMajor = 1i32,
    /// The best way to present the data is indeterminate.
    Indeterminate = 2i32
}

/// Defines enum for `windows::Win32::UI::Accessibility::ScrollAmount`.
/// 
/// Contains values that specify the direction and distance to scroll.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::ScrollAmount)]
pub enum ScrollAmount {
    /// Scrolling is done in large decrements, equivalent to pressing the PAGE UP key or clicking on a blank part of a scroll bar. 
    /// If one page up is not a relevant amount for the control and no scroll bar exists, the value represents an amount equal to the current visible window.
    LargeDecrement = 0i32,
    /// Scrolling is done in small decrements, equivalent to pressing an arrow key or clicking the arrow button on a scroll bar.
    SmallDecrement = 1i32,
    /// No scrolling is done.
    NoAmount = 2i32,
    /// Scrolling is done in large increments, equivalent to pressing the PAGE DOWN or PAGE UP key or clicking on a blank part of a scroll bar.
    /// If one page is not a relevant amount for the control and no scroll bar exists, the value represents an amount equal to the current visible window.
    LargeIncrement = 3i32,
    /// Scrolling is done in small increments, equivalent to pressing an arrow key or clicking the arrow button on a scroll bar.
    SmallIncrement = 4i32
}

/// Defines enum for `windows::Win32::UI::Accessibility::SupportedTextSelection`.
/// 
/// Contains values that specify the supported text selection attribute.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::SupportedTextSelection)]
pub enum SupportedTextSelection {
    /// Does not support text selections.
    None = 0i32,
    /// Supports a single, continuous text selection.
    Single = 1i32,
    /// Supports multiple, disjoint text selections.
    Multiple = 2i32,
}

/// Defines enum for `windows::Win32::UI::Accessibility::ToggleState`.
/// 
/// Contains values that specify the toggle state of a Microsoft UI Automation element that implements the Toggle control pattern.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::ToggleState)]
pub enum ToggleState {
    /// The UI Automation element is not selected, checked, marked or otherwise activated.
    Off = 0i32,
    /// The UI Automation element is selected, checked, marked or otherwise activated.
    On = 1i32,
    /// The UI Automation element is in an indeterminate state.
    /// 
    /// The Indeterminate property can be used to indicate whether the user has acted on a control. For example, a check box can appear checked and dimmed, indicating an indeterminate state.
    /// 
    /// Creating an indeterminate state is different from disabling the control. 
    /// Consequently, a check box in the indeterminate state can still receive the focus. 
    /// When the user clicks an indeterminate control the ToggleState cycles to its next value.
    Indeterminate = 2i32,
}

/// Defines enum for `windows::Win32::UI::Accessibility::ZoomUnit`.
/// 
/// Contains possible values for the IUIAutomationTransformPattern2::ZoomByUnit method, which zooms the viewport of a control by the specified unit.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::ZoomUnit)]
pub enum ZoomUnit {
    /// No increase or decrease in zoom.
    NoAmount = 0i32,
    /// Decrease zoom by a large decrement.
    LargeDecrement = 1i32,
    /// Decrease zoom by a small decrement.
    SmallDecrement = 2i32,
    /// Increase zoom by a large increment.
    LargeIncrement = 3i32,
    /// Increase zoom by a small increment.
    SmallIncrement = 4i32,
}

/// Defines enum for `windows::Win32::UI::Accessibility::WindowVisualState`.
/// 
/// Contains values that specify the visual state of a window for the IWindowProvider pattern.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::WindowVisualState)]
pub enum WindowVisualState {
    /// Specifies that the window is normal (restored).
    Normal = 0i32,
    /// Specifies that the window is maximized.
    Maximized = 1i32,
    /// Specifies that the window is minimized.
    Minimized = 2i32
}

/// Defines enum for `windows::Win32::UI::Accessibility::TextUnit`.
/// 
/// Contains values that specify units of text for the purposes of navigation.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::TextUnit)]
pub enum TextUnit {
    /// Specifies that the text unit is one character in length.
    Character = 0i32,
    /// Specifies that the text unit is the length of a single, common format specification, such as bold, italic, or similar.
    Format = 1i32,
    /// Specifies that the text unit is one word in length.
    Word = 2i32,
    /// Specifies that the text unit is one line in length.
    Line = 3i32,
    /// Specifies that the text unit is one paragraph in length.
    Paragraph = 4i32,
    /// Specifies that the text unit is one document-specific page in length.
    Page = 5i32,
    /// Specifies that the text unit is an entire document in length.
    Document = 6i32,
}

/// Defines enum for `windows::Win32::UI::Accessibility::TextPatternRangeEndpoint`.
/// 
/// Contains values that specify the endpoints of a text range.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::TextPatternRangeEndpoint)]
pub enum TextPatternRangeEndpoint {
    /// The starting endpoint of the range.
    Start = 0i32,
    /// The ending endpoint of the range.
    End = 1i32,
}

/// Defines enum for `windows::Win32::UI::Accessibility::OrientationType`.
/// 
/// Contains values that specify the orientation of a control.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::OrientationType)]
pub enum OrientationType {
    /// The control has no orientation.
    None = 0i32,
    /// The control has horizontal orientation.
    Horizontal = 1i32,
    /// The control has vertical orientation.
    Vertical = 2i32
}

/// Defines enum for `windows::Win32::UI::Accessibility::PropertyConditionFlags`.
/// 
/// Contains values used in creating property conditions.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::PropertyConditionFlags)]
pub enum PropertyConditionFlags {
    /// No flags.
    None = 0i32,
    /// Comparison of string properties is not case-sensitive.
    IgnoreCase = 1i32,
    /// Comparison of substring properties is enabled.
    MatchSubstring = 2i32,
    /// Combines `IgnoreCase` and `MatchSubstring` flags.
    All = 3i32
}

/// Defines enum for `windows::Win32::UI::Accessibility::TreeScope`.
/// 
/// Contains values that specify the scope of various operations in the Microsoft UI Automation tree.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::TreeScope)]
pub enum TreeScope {
    /// The scope excludes the subtree from the search.
    None = 0i32,
    /// The scope includes the element itself.
    Element = 1i32,
    /// The scope includes children of the element.
    Children = 2i32,
    /// The scope includes children and more distant descendants of the element.
    Descendants = 4i32,
    /// The scope includes the parent of the element.
    Parent = 8i32,
    /// The scope includes the parent and more distant ancestors of the element.
    Ancestors = 16i32,
    /// The scope includes the element and all its descendants. This flag is a combination of the TreeScope_Element and TreeScope_Descendants values.
    Subtree = 7i32
}

/// Defines enum for `windows::Win32::UI::Accessibility::UIA_ANNOTATIONTYPE`.
/// 
/// This type describes the named constants that are used to identify types of annotations in a document.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::UIA_ANNOTATIONTYPE)]
pub enum AnnotationType {
    /// The annotation type is unknown.
    Unknown = 60000i32,
    /// A spelling error, often denoted by a red squiggly line.
    SpellingError = 60001i32,
    /// A grammatical error, often denoted by a green squiggly line.
    GrammarError = 60002i32,
    /// A comment. Comments can take different forms depending on the application.
    Comment = 60003i32,
    /// An error in a formula. Formula errors typically include red text and exclamation marks.
    FormulaError = 60004i32,
    /// A change that was made to the document.
    TrackChanges = 60005i32,
    /// The header for a page in a document.
    Header = 60006i32,
    /// The footer for a page in a document.
    Footer = 60007i32,
    /// Highlighted content, typically denoted by a contrasting background color.
    Highlighted = 60008i32,
    /// The endnote for a document.
    Endnote = 60009i32,
    /// The footnote for a page in a document.
    Footnote = 60010i32,
    /// An insertion change that was made to the document.
    InsertionChange = 60011i32,
    /// A deletion change that was made to the document.
    DeletionChange = 60012i32,
    /// A move change that was made to the document.
    MoveChange = 60013i32,
    /// A format change that was made.
    FormatChange = 60014i32,
    /// An unsynced change that was made to the document.
    UnsyncedChange = 60015i32,
    /// An editing locked change that was made to the document.
    EditingLockedChange = 60016i32,
    /// An external change that was made to the document.
    ExternalChange = 60017i32,
    /// A conflicting change that was made to the document.
    ConflictingChange = 60018i32,
    /// The author of the document.
    Author = 60019i32,
    /// An advanced proofing issue.
    AdvancedProofingIssue = 60020i32,
    /// A data validation error that occurred.
    DataValidationError = 60021i32,
    /// A circular reference error that occurred.
    CircularReferenceError = 60022i32,
    /// A text range containing mathematics.
    Mathematics = 60023i32,
}

/// Defines enum for `windows::Win32::UI::Accessibility::UIA_STYLE_ID`.
/// 
/// This set of constants describes the named constants used to identify the visual style of text in a document.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::UIA_STYLE_ID)]
pub enum StyleType {
    /// A custom style.
    Custom = 70000i32,
    /// A first level heading.
    Heading1 = 70001i32,
    /// A second level heading.
    Heading2 = 70002i32,
    /// A third level heading.
    Heading3 = 70003i32,
    /// A fourth level heading.
    Heading4 = 70004i32,
    /// A fifth level heading.
    Heading5 = 70005i32,
    /// A sixth level heading.
    Heading6 = 70006i32,
    /// A seventh level heading.
    Heading7 = 70007i32,
    /// An eighth level heading.
    Heading8 = 70008i32,
    /// A ninth level heading.
    Heading9 = 70009i32,
    /// A title.
    Title = 70010i32,
    /// A subtitle.
    Subtitle = 70011i32,
    /// Normal style.
    Normal = 70012i32,
    /// Text that is emphasized.
    Emphasis = 70013i32,
    /// A quotation.
    Quote = 70014i32,
    /// A list with bulleted items. Supported starting with Windows 8.1.
    BulletedList = 70015i32,
    /// A list with numbered items. Supported starting with Windows 8.1.
    NumberedList = 70016i32,
}

/// Defines enum for `windows::Win32::UI::Accessibility::UIA_TEXTATTRIBUTE_ID`.
/// 
/// This type describes the named constants used to identify text attributes of a Microsoft UI Automation text range.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::UIA_TEXTATTRIBUTE_ID)]
pub enum TextAttribute {
    /// Identifies the AnimationStyle text attribute, which specifies the type of animation applied to the text. This attribute is specified as a value from the AnimationStyle enumerated type.
    AnimationStyle = 40000i32,
    /// Identifies the BackgroundColor text attribute, which specifies the background color of the text. This attribute is specified as a COLORREF; a 32-bit value used to specify an RGB or RGBA color.
    BackgroundColor = 40001i32,
    /// Identifies the BulletStyle text attribute, which specifies the style of bullets used in the text range. This attribute is specified as a value from the BulletStyle enumerated type.
    BulletStyle = 40002i32,
    /// Identifies the CapStyle text attribute, which specifies the capitalization style for the text. This attribute is specified as a value from the CapStyle enumerated type.
    CapStyle = 40003i32,
    /// Identifies the Culture text attribute, which specifies the locale of the text by locale identifier (LCID).
    Culture = 40004i32,
    /// Identifies the FontName text attribute, which specifies the name of the font. Examples: Arial Black; Arial Narrow. The font name string is not localized.
    FontName = 40005i32,
    /// Identifies the FontSize text attribute, which specifies the point size of the font.
    FontSize = 40006i32,
    /// Identifies the FontWeight text attribute, which specifies the relative stroke, thickness, or boldness of the font. The FontWeight attribute is modeled after the lfWeight member of the GDI LOGFONT structure, and related standards, and can be one of the following values:
    FontWeight = 40007i32,
    /// Identifies the ForegroundColor text attribute, which specifies the foreground color of the text. This attribute is specified as a COLORREF, a 32-bit value used to specify an RGB or RGBA color.
    ForegroundColor = 40008i32,
    /// Identifies the HorizontalTextAlignment text attribute, which specifies how the text is aligned horizontally. This attribute is specified as a value from the HorizontalTextAlignmentEnum enumerated type.
    HorizontalTextAlignment = 40009i32,
    /// Identifies the IndentationFirstLine text attribute, which specifies how far, in points, to indent the first line of a paragraph.
    IndentationFirstLine = 40010i32,
    /// Identifies the IndentationLeading text attribute, which specifies the leading indentation, in points.
    IndentationLeading = 40011i32,
    /// Identifies the IndentationTrailing text attribute, which specifies the trailing indentation, in points.
    IndentationTrailing = 40012i32,
    /// Identifies the IsHidden text attribute, which indicates whether the text is hidden (TRUE) or visible (FALSE).
    IsHidden = 40013i32,
    /// Identifies the IsItalic text attribute, which indicates whether the text is italic (TRUE) or not (FALSE).
    IsItalic = 40014i32,
    /// Identifies the IsReadOnly text attribute, which indicates whether the text is read-only (TRUE) or can be modified (FALSE).
    IsReadOnly = 40015i32,
    /// Identifies the IsSubscript text attribute, which indicates whether the text is subscript (TRUE) or not (FALSE).
    IsSubscript = 40016i32,
    /// Identifies the IsSuperscript text attribute, which indicates whether the text is subscript (TRUE) or not (FALSE).
    IsSuperscript = 40017i32,
    /// Identifies the MarginBottom text attribute, which specifies the size, in points, of the bottom margin applied to the page associated with the text range.
    MarginBottom = 40018i32,
    /// Identifies the MarginLeading text attribute, which specifies the size, in points, of the leading margin applied to the page associated with the text range.
    MarginLeading = 40019i32,
    /// Identifies the MarginTop text attribute, which specifies the size, in points, of the top margin applied to the page associated with the text range.
    MarginTop = 40020i32,
    /// Identifies the MarginTrailing text attribute, which specifies the size, in points, of the trailing margin applied to the page associated with the text range.
    MarginTrailing = 40021i32,
    /// Identifies the OutlineStyles text attribute, which specifies the outline style of the text. This attribute is specified as a value from the OutlineStyles enumerated type.
    OutlineStyles = 40022i32,
    /// Identifies the OverlineColor text attribute, which specifies the color of the overline text decoration. This attribute is specified as a COLORREF, a 32-bit value used to specify an RGB or RGBA color.
    OverlineColor = 40023i32,
    /// Identifies the OverlineStyle text attribute, which specifies the style of the overline text decoration. This attribute is specified as a value from the TextDecorationLineStyleEnum enumerated type.
    OverlineStyle = 40024i32,
    /// Identifies the StrikethroughColor text attribute, which specifies the color of the strikethrough text decoration. This attribute is specified as a COLORREF, a 32-bit value used to specify an RGB or RGBA color.
    StrikethroughColor = 40025i32,
    /// Identifies the StrikethroughStyle text attribute, which specifies the style of the strikethrough text decoration. This attribute is specified as a value from the TextDecorationLineStyleEnum enumerated type.
    StrikethroughStyle = 40026i32,
    /// Identifies the Tabs text attribute, which is an array specifying the tab stops for the text range. Each array element specifies a distance, in points, from the leading margin.
    Tabs = 40027i32,
    /// Identifies the TextFlowDirections text attribute, which specifies the direction of text flow. This attribute is specified as a combination of values from the FlowDirections enumerated type.
    TextFlowDirections = 40028i32,
    /// Identifies the UnderlineColor text attribute, which specifies the color of the underline text decoration. This attribute is specified as a COLORREF, a 32-bit value used to specify an RGB or RGBA color.
    UnderlineColor = 40029i32,
    /// Identifies the UnderlineStyle text attribute, which specifies the style of the underline text decoration. This attribute is specified as a value from the TextDecorationLineStyleEnum enumerated type.
    UnderlineStyle = 40030i32,
    /// Identifies the AnnotationTypes text attribute, which maintains a list of annotation type identifiers for a range of text. For a list of possible values, see Annotation Type Identifiers. Supported starting with Windows 8.
    AnnotationTypes = 40031i32,
    /// Identifies the AnnotationObjects text attribute, which maintains an array of IUIAutomationElement2 interfaces, one for each element in the current text range that implements the Annotation control pattern. Each element might also implement other control patterns as needed to describe the annotation. For example, an annotation that is a comment would also support the Text control pattern. Supported starting with Windows 8.
    AnnotationObjects = 40032i32,
    /// Identifies the StyleName text attribute, which identifies the localized name of the text style in use for a text range. Supported starting with Windows 8.
    StyleName = 40033i32,
    /// Identifies the StyleId text attribute, which indicates the text styles in use for a text range. For a list of possible values, see Style Identifiers. Supported starting with Windows 8.
    StyleId = 40034i32,
    /// Identifies the Link text attribute, which contains the IUIAutomationTextRange interface of the text range that is the target of an internal link in a document. Supported starting with Windows 8.
    Link = 40035i32,
    /// Identifies the IsActive text attribute, which indicates whether the control that contains the text range has the keyboard focus (TRUE) or not (FALSE). Supported starting with Windows 8.
    IsActive = 40036i32,
    /// Identifies the SelectionActiveEnd text attribute, which indicates the location of the caret relative to a text range that represents the currently selected text. This attribute is specified as a value from the ActiveEnd enumeration. Supported starting with Windows 8.
    SelectionActiveEnd = 40037i32,
    /// Identifies the CaretPosition text attribute, which indicates whether the caret is at the beginning or the end of a line of text in the text range. This attribute is specified as a value from the CaretPosition enumerated type. Supported starting with Windows 8.
    CaretPosition = 40038i32,
    /// Identifies the CaretBidiMode text attribute, which indicates the direction of text flow in the text range. This attribute is specified as a value from the CaretBidiMode enumerated type. Supported starting with Windows 8.
    CaretBidiMode = 40039i32,
    /// Identifies the LineSpacing text attribute, which specifies the spacing between lines of text.
    LineSpacing = 40040i32,
    /// Identifies the BeforeParagraphSpacing text attribute, which specifies the size of spacing before the paragraph.
    BeforeParagraphSpacing = 40041i32,
    /// Identifies the AfterParagraphSpacing text attribute, which specifies the size of spacing after the paragraph.
    AfterParagraphSpacing = 40042i32,
}

/// Defines enum for `windows::Win32::UI::Accessibility::AutomationElementMode`.
/// 
/// Contains values that specify the type of reference to use when returning UI Automation elements.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::AutomationElementMode)]
pub enum ElementMode {
    /// Specifies that returned elements have no reference to the underlying UI and contain only cached information.
    None = 0i32,
    /// Specifies that returned elements have a full reference to the underlying UI.
    Full = 1i32
}

/// `StructureChangeType` is an enum wrapper for `windows::Win32::UI::Accessibility::StructureChangeType`.
/// 
/// Contains values that specify the type of change in the Microsoft UI Automation tree structure.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
#[map_as(windows::Win32::UI::Accessibility::StructureChangeType)]
pub enum StructureChangeType {
    /// A child element was added to the UI Automation element tree.
    ChildAdded = 0i32,
    /// A child element was removed from the UI Automation element tree.
    ChildRemoved = 1i32,
    /// Child elements were invalidated in the UI Automation element tree. 
    /// This might mean that one or more child elements were added or removed, or a combination of both. 
    /// This value can also indicate that one subtree in the UI was substituted for another. 
    /// For example, the entire contents of a dialog box changed at once, or the view of a list changed because an Explorer-type application navigated to another location. 
    /// The exact meaning depends on the UI Automation provider implementation.
    ChildrenInvalidated = 2i32,
    /// Child elements were added in bulk to the UI Automation element tree.
    ChildrenBulkAdded = 3i32,
    /// Child elements were removed in bulk from the UI Automation element tree.
    ChildrenBulkRemoved = 4i32,
    /// The order of child elements has changed in the UI Automation element tree. Child elements may or may not have been added or removed.
    ChildrenReordered = 5i32
}

#[cfg(test)]
mod tests {
    use windows::Win32::UI::Accessibility;

    use super::WindowInteractionState;

    #[test]
    fn test_window_interaction_state() {
        assert_eq!(Ok(WindowInteractionState::Running), WindowInteractionState::try_from(0));
        assert_eq!(Ok(WindowInteractionState::NotResponding), WindowInteractionState::try_from(4));
        assert!(WindowInteractionState::try_from(100).is_err());
        
        assert_eq!(1i32, WindowInteractionState::Closing as i32);

        assert_eq!(Accessibility::WindowInteractionState_ReadyForUserInteraction, WindowInteractionState::ReadyForUserInteraction.into());
        assert_eq!(WindowInteractionState::Running, Accessibility::WindowInteractionState_Running.into());

        let running = format!("{}", WindowInteractionState::Running);
        assert_eq!(running, "Running");
    }

    #[test]
    fn test_handle() {
        let handle = crate::types::Handle::from(0x001);
        assert_eq!(windows::Win32::Foundation::HWND(unsafe { std::mem::transmute(0x001i64) } ), handle.into());
    }
}