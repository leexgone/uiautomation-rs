use std::fmt::Debug;
use std::fmt::Display;

use uiautomation_derive::EnumConvert;
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

/// Defines enum for `UIA_PROPERTY_ID`.
/// 
/// Describes the named constants that identify the properties of Microsoft UI Automation elements.
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
pub enum UIProperty {
    /// Identifies the RuntimeId property, which is an array of integers representing the identifier for an automation element.
    RuntimeId = 30000u32,
    /// Identifies the BoundingRectangle property, which specifies the coordinates of the rectangle that completely encloses the automation element.
    BoundingRectangle = 30001u32,
    /// Identifies the ProcessId property, which is an integer representing the process identifier (ID) of the automation element.
    ProcessId = 30002u32,
    /// Identifies the ControlType property, which is a class that identifies the type of the automation element. 
    ControlType = 30003u32,
    /// Identifies the LocalizedControlType property, which is a text string describing the type of control that the automation element represents.
    LocalizedControlType = 30004u32,
    /// Identifies the Name property, which is a string that holds the name of the automation element.
    Name = 30005u32,
    /// Identifies the AcceleratorKey property, which is a string containing the accelerator key (also called shortcut key) combinations for the automation element.
    AcceleratorKey = 30006u32,
    /// Identifies the AccessKey property, which is a string containing the access key character for the automation element.
    AccessKey = 30007u32,
    /// Identifies the HasKeyboardFocus property, which is a Boolean value that indicates whether the automation element has keyboard focus.
    HasKeyboardFocus = 30008u32,
    /// Identifies the IsKeyboardFocusable property, which is a Boolean value that indicates whether the automation element can accept keyboard focus.
    IsKeyboardFocusable = 30009u32,
    /// Identifies the IsEnabled property, which is a Boolean value that indicates whether the UI item referenced by the automation element is enabled and can be interacted with.
    IsEnabled = 30010u32,
    /// Identifies the AutomationId property, which is a string containing the UI Automation identifier (ID) for the automation element.
    AutomationId = 30011u32,
    /// Identifies the ClassName property, which is a string containing the class name for the automation element as assigned by the control developer.
    ClassName = 30012u32,
    /// Identifies the HelpText property, which is a help text string associated with the automation element.
    HelpText = 30013u32,
    /// Identifies the ClickablePoint property, which is a point on the automation element that can be clicked. 
    ClickablePoint = 30014u32,
    /// Identifies the Culture property, which contains a locale identifier for the automation element.
    Culture = 30015u32,
    /// Identifies the IsControlElement property, which is a Boolean value that specifies whether the element appears in the control view of the automation element tree.
    IsControlElement = 30016u32,
    /// Identifies the IsContentElement property, which is a Boolean value that specifies whether the element appears in the content view of the automation element tree.
    IsContentElement = 30017u32,
    /// Identifies the LabeledBy property, which is an automation element that contains the text label for this element.
    LabeledBy = 30018u32,
    /// Identifies the IsPassword property, which is a Boolean value that indicates whether the automation element contains protected content or a password.
    IsPassword = 30019u32,
    /// Identifies the NativeWindowHandle property, which is an integer that represents the handle (HWND) of the automation element window, if it exists; otherwise, this property is 0.
    NativeWindowHandle = 30020u32,
    /// Identifies the ItemType property, which is a text string describing the type of the automation element.
    ItemType = 30021u32,
    /// Identifies the IsOffscreen property, which is a Boolean value that indicates whether the automation element is entirely scrolled out of view (for example, an item in a list box that is outside the viewport of the container object) or collapsed out of view (for example, an item in a tree view or menu, or in a minimized window). 
    IsOffscreen = 30022u32,
    /// Identifies the Orientation property, which indicates the orientation of the control represented by the automation element. The property is expressed as a value from the OrientationType enumerated type.
    Orientation = 30023u32,
    /// Identifies the FrameworkId property, which is a string containing the name of the underlying UI framework that the automation element belongs to.
    FrameworkId = 30024u32,
    /// Identifies the IsRequiredForForm property, which is a Boolean value that indicates whether the automation element is required to be filled out on a form.
    IsRequiredForForm = 30025u32,
    /// Identifies the ItemStatus property, which is a text string describing the status of an item of the automation element.
    ItemStatus = 30026u32,
    /// Identifies the IsDockPatternAvailable property, which indicates whether the Dock control pattern is available for the automation element.
    IsDockPatternAvailable = 30027u32,
    /// Identifies the IsExpandCollapsePatternAvailable property, which indicates whether the ExpandCollapse control pattern is available for the automation element.
    IsExpandCollapsePatternAvailable = 30028u32,
    /// Identifies the IsGridItemPatternAvailable property, which indicates whether the GridItem control pattern is available for the automation element.
    IsGridItemPatternAvailable = 30029u32,
    /// Identifies the IsGridPatternAvailable property, which indicates whether the Grid control pattern is available for the automation element.
    IsGridPatternAvailable = 30030u32,
    /// Identifies the IsInvokePatternAvailable property, which indicates whether the Invoke control pattern is available for the automation element.
    IsInvokePatternAvailable = 30031u32,
    /// Identifies the IsMultipleViewPatternAvailable property, which indicates whether the MultipleView control pattern is available for the automation element.
    IsMultipleViewPatternAvailable = 30032u32,
    /// Identifies the IsRangeValuePatternAvailable property, which indicates whether the RangeValue control pattern is available for the automation element.
    IsRangeValuePatternAvailable = 30033u32,
    /// Identifies the IsScrollPatternAvailable property, which indicates whether the Scroll control pattern is available for the automation element.
    IsScrollPatternAvailable = 30034u32,
    /// Identifies the IsScrollItemPatternAvailable property, which indicates whether the ScrollItem control pattern is available for the automation element.
    IsScrollItemPatternAvailable = 30035u32,
    /// Identifies the IsSelectionItemPatternAvailable property, which indicates whether the SelectionItem control pattern is available for the automation element. 
    IsSelectionItemPatternAvailable = 30036u32,
    /// Identifies the IsSelectionPatternAvailable property, which indicates whether the Selection control pattern is available for the automation element.
    IsSelectionPatternAvailable = 30037u32,
    /// Identifies the IsTablePatternAvailable property, which indicates whether the Table control pattern is available for the automation element.
    IsTablePatternAvailable = 30038u32,
    /// Identifies the IsTableItemPatternAvailable property, which indicates whether the TableItem control pattern is available for the automation element.
    IsTableItemPatternAvailable = 30039u32,
    /// Identifies the IsTextPatternAvailable property, which indicates whether the Text control pattern is available for the automation element.
    IsTextPatternAvailable = 30040u32,
    /// Identifies the IsTogglePatternAvailable property, which indicates whether the Toggle control pattern is available for the automation element.
    IsTogglePatternAvailable = 30041u32,
    /// Identifies the IsTransformPatternAvailable property, which indicates whether the Transform control pattern is available for the automation element.
    IsTransformPatternAvailable = 30042u32,
    /// Identifies the IsValuePatternAvailable property, which indicates whether the Value control pattern is available for the automation element.
    IsValuePatternAvailable = 30043u32,
    /// Identifies the IsWindowPatternAvailable property, which indicates whether the Window control pattern is available for the automation element.
    IsWindowPatternAvailable = 30044u32,
    /// Identifies the Value property of the Value control pattern.
    ValueValue = 30045u32,
    /// Identifies the IsReadOnly property of the Value control pattern.
    ValueIsReadOnly = 30046u32,
    /// Identifies the Value property of the RangeValue control pattern.
    RangeValueValue = 30047u32,
    /// Identifies the IsReadOnly property of the RangeValue control pattern.
    RangeValueIsReadOnly = 30048u32,
    /// Identifies the Minimum property of the RangeValue control pattern.
    RangeValueMinimum = 30049u32,
    /// Identifies the Maximum property of the RangeValue control pattern.
    RangeValueMaximum = 30050u32,
    /// Identifies the LargeChange property of the RangeValue control pattern.
    RangeValueLargeChange = 30051u32,
    /// Identifies the SmallChange property of the RangeValue control pattern.
    RangeValueSmallChange = 30052u32,
    /// Identifies the HorizontalScrollPercent property of the Scroll control pattern.
    ScrollHorizontalScrollPercent = 30053u32,
    /// Identifies the HorizontalViewSize property of the Scroll control pattern.
    ScrollHorizontalViewSize = 30054u32,
    /// Identifies the VerticalScrollPercent property of the Scroll control pattern.
    ScrollVerticalScrollPercent = 30055u32,
    /// Identifies the VerticalViewSize property of the Scroll control pattern.
    ScrollVerticalViewSize = 30056u32,
    /// Identifies the HorizontallyScrollable property of the Scroll control pattern.
    ScrollHorizontallyScrollable = 30057u32,
    /// Identifies the VerticallyScrollable property of the Scroll control pattern.
    ScrollVerticallyScrollable = 30058u32,
    /// Identifies the Selection property of the Selection control pattern.
    SelectionSelection = 30059u32,
    /// Identifies the CanSelectMultiple property of the Selection control pattern.
    SelectionCanSelectMultiple = 30060u32,
    /// Identifies the IsSelectionRequired property of the Selection control pattern.
    SelectionIsSelectionRequired = 30061u32,
    /// Identifies the RowCount property of the Grid control pattern. This property indicates the total number of rows in the grid.
    GridRowCount = 30062u32,
    /// Identifies the ColumnCount property of the Grid control pattern. This property indicates the total number of columns in the grid.
    GridColumnCount = 30063u32,
    /// Identifies the Row property of the GridItem control pattern. This property is the ordinal number of the row that contains the cell or item.
    GridItemRow = 30064u32,
    /// Identifies the Column property of the GridItem control pattern. This property indicates the ordinal number of the column that contains the cell or item.
    GridItemColumn = 30065u32,
    /// Identifies the RowSpan property of the GridItem control pattern. This property indicates the number of rows spanned by the cell or item.
    GridItemRowSpan = 30066u32,
    /// Identifies the ColumnSpan property of the GridItem control pattern. This property indicates the number of columns spanned by the cell or item.
    GridItemColumnSpan = 30067u32,
    /// Identifies the ContainingGrid property of the GridItem control pattern.
    GridItemContainingGrid = 30068u32,
    /// Identifies the DockPosition property of the Dock control pattern.
    DockDockPosition = 30069u32,
    /// Identifies the ExpandCollapseState property of the ExpandCollapse control pattern.
    ExpandCollapseExpandCollapseState = 30070u32,
    /// Identifies the CurrentView property of the MultipleView control pattern.
    MultipleViewCurrentView = 30071u32,
    /// Identifies the SupportedViews property of the MultipleView control pattern.
    MultipleViewSupportedViews = 30072u32,
    /// Identifies the CanMaximize property of the Window control pattern.
    WindowCanMaximize = 30073u32,
    /// Identifies the CanMinimize property of the Window control pattern.
    WindowCanMinimize = 30074u32,
    /// Identifies the WindowVisualState property of the Window control pattern.
    WindowWindowVisualState = 30075u32,
    /// Identifies the WindowInteractionState property of the Window control pattern.
    WindowWindowInteractionState = 30076u32,
    /// Identifies the IsModal property of the Window control pattern.
    WindowIsModal = 30077u32,
    /// Identifies the IsTopmost property of the Window control pattern.
    WindowIsTopmost = 30078u32,
    /// Identifies the IsSelected property of the SelectionItem control pattern.
    SelectionItemIsSelected = 30079u32,
    /// Identifies the SelectionContainer property of the SelectionItem control pattern.
    SelectionItemSelectionContainer = 30080u32,
    /// Identifies the RowHeaders property of the Table control pattern.
    TableRowHeaders = 30081u32,
    /// Identifies the ColumnHeaders property of the Table control pattern.
    TableColumnHeaders = 30082u32,
    /// Identifies the RowOrColumnMajor property of the Table control pattern.
    TableRowOrColumnMajor = 30083u32,
    /// Identifies the RowHeaderItems property of the TableItem control pattern.
    TableItemRowHeaderItems = 30084u32,
    /// Identifies the ColumnHeaderItems property of the TableItem control pattern.
    TableItemColumnHeaderItems = 30085u32,
    /// Identifies the ToggleState property of the Toggle control pattern.
    ToggleToggleState = 30086u32,
    /// Identifies the CanMove property of the Transform control pattern.
    TransformCanMove = 30087u32,
    /// Identifies the CanResize property of the Transform control pattern.
    TransformCanResize = 30088u32,
    /// Identifies the CanRotate property of the Transform control pattern.
    TransformCanRotate = 30089u32,
    /// Identifies the IsLegacyIAccessiblePatternAvailable property, which indicates whether the LegacyIAccessible control pattern is available for the automation element.
    IsLegacyIAccessiblePatternAvailable = 30090u32,
    /// Identifies the ChildId property of the LegacyIAccessible control pattern.
    LegacyIAccessibleChildId = 30091u32,
    /// Identifies the Name property of the LegacyIAccessible control pattern.
    LegacyIAccessibleName = 30092u32,
    /// Identifies the Value property of the LegacyIAccessible control pattern.
    LegacyIAccessibleValue = 30093u32,
    /// Identifies the Description property of the LegacyIAccessible control pattern.
    LegacyIAccessibleDescription = 30094u32,
    /// Identifies the Roleproperty of the LegacyIAccessible control pattern.
    LegacyIAccessibleRole = 30095u32,
    /// Identifies the State property of the LegacyIAccessible control pattern.
    LegacyIAccessibleState = 30096u32,
    /// Identifies the Help property of the LegacyIAccessible control pattern.
    LegacyIAccessibleHelp = 30097u32,
    /// Identifies the KeyboardShortcut property of the LegacyIAccessible control pattern.
    LegacyIAccessibleKeyboardShortcut = 30098u32,
    /// Identifies the Selection property of the LegacyIAccessible control pattern.
    LegacyIAccessibleSelection = 30099u32,
    /// Identifies the DefaultAction property of the LegacyIAccessible control pattern.
    LegacyIAccessibleDefaultAction = 30100u32,
    /// Identifies the AriaRole property, which is a string containing the Accessible Rich Internet Application (ARIA) role information for the automation element. 
    AriaRole = 30101u32,
    /// Identifies the AriaProperties property, which is a formatted string containing the Accessible Rich Internet Application (ARIA) property information for the automation element. 
    AriaProperties = 30102u32,
    /// Identifies the IsDataValidForForm property, which is a Boolean value that indicates whether the entered or selected value is valid for the form rule associated with the automation element. 
    IsDataValidForForm = 30103u32,
    /// Identifies the ControllerFor property, which is an array of automation elements that are manipulated by the automation element that supports this property.
    ControllerFor = 30104u32,
    /// Identifies the DescribedBy property, which is an array of elements that provide more information about the automation element.
    DescribedBy = 30105u32,
    /// Identifies the FlowsTo property, which is an array of automation elements that suggests the reading order after the current automation element.
    FlowsTo = 30106u32,
    /// Identifies the ProviderDescription property, which is a formatted string containing the source information of the UI Automation provider for the automation element, including proxy information.
    ProviderDescription = 30107u32,
    /// Identifies the IsItemContainerPatternAvailable property, which indicates whether the ItemContainer control pattern is available for the automation element.
    IsItemContainerPatternAvailable = 30108u32,
    /// Identifies the IsVirtualizedItemPatternAvailable property, which indicates whether the VirtualizedItem control pattern is available for the automation element.
    IsVirtualizedItemPatternAvailable = 30109u32,
    /// Identifies the IsSynchronizedInputPatternAvailable property, which indicates whether the SynchronizedInput control pattern is available for the automation element.
    IsSynchronizedInputPatternAvailable = 30110u32,
    /// Identifies the OptimizeForVisualContent property, which is a Boolean value that indicates whether the provider exposes only elements that are visible.
    OptimizeForVisualContent = 30111u32,
    /// Identifies the IsObjectModelPatternAvailable property, which indicates whether the ObjectModel control pattern is available for the automation element.
    IsObjectModelPatternAvailable = 30112u32,
    /// Identifies the AnnotationTypeId property of the Annotation control pattern. Supported starting with Windows 8.
    AnnotationAnnotationTypeId = 30113u32,
    /// Identifies the AnnotationTypeName property of the Annotation control pattern. Supported starting with Windows 8.
    AnnotationAnnotationTypeName = 30114u32,
    /// Identifies the Author property of the Annotation control pattern. Supported starting with Windows 8.
    AnnotationAuthor = 30115u32,
    /// Identifies the DateTime property of the Annotation control pattern. Supported starting with Windows 8.
    AnnotationDateTime = 30116u32,
    /// Identifies the Target property of the Annotation control pattern. Supported starting with Windows 8.
    AnnotationTarget = 30117u32,
    /// Identifies the IsAnnotationPatternAvailable property, which indicates whether the Annotation control pattern is available for the automation element.
    IsAnnotationPatternAvailable = 30118u32,
    /// Identifies the IsTextPattern2Available property, which indicates whether version two of the Text control pattern is available for the automation element.
    IsTextPattern2Available = 30119u32,
    /// Identifies the StyleId property of the Styles control pattern.
    StylesStyleId = 30120u32,
    /// Identifies the StyleName property of the Styles control pattern.
    StylesStyleName = 30121u32,
    /// Identifies the FillColor property of the Styles control pattern.
    StylesFillColor = 30122u32,
    /// Identifies the FillPatternStyle property of the Styles control pattern.
    StylesFillPatternStyle = 30123u32,
    /// Identifies the Shape property of the Styles control pattern.
    StylesShape = 30124u32,
    /// Identifies the FillPatternColor property of the Styles control pattern.
    StylesFillPatternColor = 30125u32,
    /// Identifies the ExtendedProperties property of the Styles control pattern.
    StylesExtendedProperties = 30126u32,
    /// Identifies the IsStylesPatternAvailable property, which indicates whether the Styles control pattern is available for the automation element.
    IsStylesPatternAvailable = 30127u32,
    /// Identifies the IsSpreadsheetPatternAvailable property, which indicates whether the Spreadsheet control pattern is available for the automation element.
    IsSpreadsheetPatternAvailable = 30128u32,
    /// Identifies the Formula property of the SpreadsheetItem control pattern.
    SpreadsheetItemFormula = 30129u32,
    /// Identifies the AnnotationObjects property of the SpreadsheetItem control pattern.
    SpreadsheetItemAnnotationObjects = 30130u32,
    /// Identifies the AnnotationTypes property of the SpreadsheetItem control pattern. Supported starting with Windows 8.
    SpreadsheetItemAnnotationTypes = 30131u32,
    /// Identifies the IsSpreadsheetItemPatternAvailable property, which indicates whether the SpreadsheetItem control pattern is available for the automation element.
    IsSpreadsheetItemPatternAvailable = 30132u32,
    /// Identifies the CanZoom property of the Transform control pattern. Supported starting with Windows 8.
    Transform2CanZoom = 30133u32,
    /// Identifies the IsTransformPattern2Available property, which indicates whether version two of the Transform control pattern is available for the automation element.
    IsTransformPattern2Available = 30134u32,
    /// Identifies the LiveSetting property, which is supported by an automation element that represents a live region.
    LiveSetting = 30135u32,
    /// Identifies the IsTextChildPatternAvailable property, which indicates whether the TextChild control pattern is available for the automation element.
    IsTextChildPatternAvailable = 30136u32,
    /// Identifies the IsDragPatternAvailable property, which indicates whether the Drag control pattern is available for the automation element.
    IsDragPatternAvailable = 30137u32,
    /// Identifies the IsGrabbed property of the Drag control pattern. Supported starting with Windows 8.
    DragIsGrabbed = 30138u32,
    /// Identifies the DropEffect property of the Drag control pattern. Supported starting with Windows 8.
    DragDropEffect = 30139u32,
    /// Identifies the DropEffects property of the Drag control pattern. Supported starting with Windows 8.
    DragDropEffects = 30140u32,
    /// Identifies the IsDropTargetPatternAvailable property, which indicates whether the DropTarget control pattern is available for the automation element.
    IsDropTargetPatternAvailable = 30141u32,
    /// Identifies the DropTargetEffect property of the DropTarget control pattern. Supported starting with Windows 8.
    DropTargetDropTargetEffect = 30142u32,
    /// Identifies the DropTargetEffects property of the DropTarget control pattern. Supported starting with Windows 8.
    DropTargetDropTargetEffects = 30143u32,
    /// Identifies the GrabbedItems property of the Drag control pattern. Supported starting with Windows 8.
    DragGrabbedItems = 30144u32,
    /// Identifies the ZoomLevel property of the Transform control pattern. Supported starting with Windows 8.
    Transform2ZoomLevel = 30145u32,
    /// Identifies the ZoomMinimum property of the Transform control pattern. Supported starting with Windows 8.
    Transform2ZoomMinimum = 30146u32,
    /// Identifies the ZoomMaximum property of the Transform control pattern. Supported starting with Windows 8.
    Transform2ZoomMaximum = 30147u32,
    /// Identifies the FlowsFrom property, which is an array of automation elements that suggests the reading order before the current automation element. Supported starting with Windows 8.
    FlowsFrom = 30148u32,
    /// Identifies the IsTextEditPatternAvailable property, which indicates whether the TextEdit control pattern is available for the automation element.
    IsTextEditPatternAvailable = 30149u32,
    /// Identifies the IsPeripheral property, which is a Boolean value that indicates whether the automation element represents peripheral UI.
    IsPeripheral = 30150u32,
    /// Identifies the IsCustomNavigationPatternAvailable property, which indicates whether the CustomNavigation control pattern is available for the automation element.
    IsCustomNavigationPatternAvailable = 30151u32,
    /// Identifies the PositionInSet property, which is a 1-based integer associated with an automation element.
    PositionInSet = 30152u32,
    /// Identifies the SizeOfSet property, which is a 1-based integer associated with an automation element.
    SizeOfSet = 30153u32,
    /// Identifies the Level property, which is a 1-based integer associated with an automation element.
    Level = 30154u32,
    /// Identifies the AnnotationTypes property, which is a list of the types of annotations in a document, such as comment, header, footer, and so on.
    AnnotationTypes = 30155u32,
    /// Identifies the AnnotationObjects property, which is a list of annotation objects in a document, such as comment, header, footer, and so on.
    AnnotationObjects = 30156u32,
    /// Identifies the LandmarkType property, which is a Landmark Type Identifier associated with an element.
    LandmarkType = 30157u32,
    /// Identifies the LocalizedLandmarkType, which is a text string describing the type of landmark that the automation element represents.
    LocalizedLandmarkType = 30158u32,
    /// The FullDescription property exposes a localized string which can contain extended description text for an element. 
    FullDescription = 30159u32,
    /// Identifies the FillColor property, which specifies the color used to fill the automation element.
    FillColor = 30160u32,
    /// Identifies the OutlineColor property, which specifies the color used for the outline of the automation element.
    OutlineColor = 30161u32,
    /// Identifies the FillType property, which specifies the pattern used to fill the automation element, such as none, color, gradient, picture, pattern, and so on.
    FillType = 30162u32,
    /// Identifies the VisualEffects property, which is a bit field that specifies effects on the automation element, such as shadow, reflection, glow, soft edges, or bevel.
    VisualEffects = 30163u32,
    /// Identifies the OutlineThickness property, which specifies the width for the outline of the automation element.
    OutlineThickness = 30164u32,
    /// Identifies the CenterPoint property, which specifies the center X and Y point coordinates of the automation element. 
    CenterPoint = 30165u32,
    /// Identifies the Rotation property, which specifies the angle of rotation in unspecified units.
    Rotation = 30166u32,
    /// Identifies the Size property, which specifies the width and height of the automation element.
    Size = 30167u32,
    /// Identifies whether the Selection2 control pattern is available.
    IsSelectionPattern2Available = 30168u32,
    /// Identifies the FirstSelectedItem property of the Selection2 control pattern.
    Selection2FirstSelectedItem = 30169u32,
    /// Identifies the LastSelectedItem property of the Selection2 control pattern.
    Selection2LastSelectedItem = 30170u32,
    /// Identifies the CurrentSelectedItem property of the Selection2 control pattern.
    Selection2CurrentSelectedItem = 30171u32,
    /// Identifies the ItemCount property of the Selection2 control pattern.
    Selection2ItemCount = 30172u32,
    /// Identifies the HeadingLevel property, which indicates the heading level of a UI Automation element.
    HeadingLevel = 30173u32,
    /// Identifies the IsDialog property, which is a Boolean value that indicates whether the automation element is a dialog window. 
    IsDialog = 30174u32,
}

impl From<windows::Win32::UI::Accessibility::UIA_PROPERTY_ID> for UIProperty {
    fn from(prop: windows::Win32::UI::Accessibility::UIA_PROPERTY_ID) -> Self {
        prop.0.try_into().unwrap()
    }
}

impl Into<windows::Win32::UI::Accessibility::UIA_PROPERTY_ID> for UIProperty {
    fn into(self) -> windows::Win32::UI::Accessibility::UIA_PROPERTY_ID {
        windows::Win32::UI::Accessibility::UIA_PROPERTY_ID(self as u32)
    }
}

/// Defines enum for `windows::Win32::UI::Accessibility::WindowInteractionState`.
/// 
/// Contains values that specify the current state of the window for purposes of user interaction.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
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

impl From<windows::Win32::UI::Accessibility::WindowInteractionState> for WindowInteractionState {
    fn from(state: windows::Win32::UI::Accessibility::WindowInteractionState) -> Self {
        state.0.try_into().unwrap()
    }
}

impl Into<windows::Win32::UI::Accessibility::WindowInteractionState> for WindowInteractionState {
    fn into(self) -> windows::Win32::UI::Accessibility::WindowInteractionState {
        windows::Win32::UI::Accessibility::WindowInteractionState(self as i32)
    }
}

/// Defines enum for `windows::Win32::UI::Accessibility::DockPosition`.
/// 
/// Contains values that specify the dock position of an object, represented by a DockPattern, within a docking container.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
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

impl From<windows::Win32::UI::Accessibility::DockPosition> for DockPosition {
    fn from(value: windows::Win32::UI::Accessibility::DockPosition) -> Self {
        value.0.try_into().unwrap()
    }
}

impl Into<windows::Win32::UI::Accessibility::DockPosition> for DockPosition {
    fn into(self) -> windows::Win32::UI::Accessibility::DockPosition {
        windows::Win32::UI::Accessibility::DockPosition(self as _)
    }
}

/// Defines enum for `windows::Win32::UI::Accessibility::ExpandCollapseState`.
/// 
/// Contains values that specify the ExpandCollapseState automation property value of a UI Automation element.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
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

impl From<windows::Win32::UI::Accessibility::ExpandCollapseState> for ExpandCollapseState {
    fn from(value: windows::Win32::UI::Accessibility::ExpandCollapseState) -> Self {
        value.0.try_into().unwrap()
    }
}

impl Into<windows::Win32::UI::Accessibility::ExpandCollapseState> for ExpandCollapseState {
    fn into(self) -> windows::Win32::UI::Accessibility::ExpandCollapseState {
        windows::Win32::UI::Accessibility::ExpandCollapseState(self as _)
    }
}

/// Defines enum for `windows::Win32::UI::Accessibility::NavigateDirection`.
/// 
/// Contains values used to specify the direction of navigation within the Microsoft UI Automation tree.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
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

impl From<windows::Win32::UI::Accessibility::NavigateDirection> for NavigateDirection {
    fn from(value: windows::Win32::UI::Accessibility::NavigateDirection) -> Self {
        value.try_into().unwrap()
    }
}

impl Into<windows::Win32::UI::Accessibility::NavigateDirection> for NavigateDirection {
    fn into(self) -> windows::Win32::UI::Accessibility::NavigateDirection {
        windows::Win32::UI::Accessibility::NavigateDirection(self as _)
    }
}

/// Defines enum for `windows::Win32::UI::Accessibility::RowOrColumnMajor`.
/// 
/// Contains values that specify whether data in a table should be read primarily by row or by column.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
pub enum RowOrColumnMajor {
    /// Data in the table should be read row by row.
    RowMajor = 0i32,
    /// Data in the table should be read column by column.
    ColumnMajor = 1i32,
    /// The best way to present the data is indeterminate.
    Indeterminate = 2i32
}

impl From<windows::Win32::UI::Accessibility::RowOrColumnMajor> for RowOrColumnMajor {
    fn from(value: windows::Win32::UI::Accessibility::RowOrColumnMajor) -> Self {
        value.0.try_into().unwrap()
    }
}

impl Into<windows::Win32::UI::Accessibility::RowOrColumnMajor> for RowOrColumnMajor {
    fn into(self) -> windows::Win32::UI::Accessibility::RowOrColumnMajor {
        windows::Win32::UI::Accessibility::RowOrColumnMajor(self as _)
    }
}

/// Defines enum for `windows::Win32::UI::Accessibility::ScrollAmount`.
/// 
/// Contains values that specify the direction and distance to scroll.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
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

impl From<windows::Win32::UI::Accessibility::ScrollAmount> for ScrollAmount {
    fn from(value: windows::Win32::UI::Accessibility::ScrollAmount) -> Self {
        value.0.try_into().unwrap()
    }
}

impl Into<windows::Win32::UI::Accessibility::ScrollAmount> for ScrollAmount {
    fn into(self) -> windows::Win32::UI::Accessibility::ScrollAmount {
        windows::Win32::UI::Accessibility::ScrollAmount(self as _)
    }
}

/// Defines enum for `windows::Win32::UI::Accessibility::SupportedTextSelection`.
/// 
/// Contains values that specify the supported text selection attribute.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, EnumConvert)]
pub enum SupportedTextSelection {
    /// Does not support text selections.
    None = 0i32,
    /// Supports a single, continuous text selection.
    Single = 1i32,
    /// Supports multiple, disjoint text selections.
    Multiple = 2i32,
}

impl From<windows::Win32::UI::Accessibility::SupportedTextSelection> for SupportedTextSelection {
    fn from(value: windows::Win32::UI::Accessibility::SupportedTextSelection) -> Self {
        value.0.try_into().unwrap()
    }
}

impl Into<windows::Win32::UI::Accessibility::SupportedTextSelection> for SupportedTextSelection {
    fn into(self) -> windows::Win32::UI::Accessibility::SupportedTextSelection {
        windows::Win32::UI::Accessibility::SupportedTextSelection(self as _)
    }
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
    }
}