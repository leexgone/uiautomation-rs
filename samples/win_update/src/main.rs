use uiautomation::UIAutomation;
use uiautomation::actions::Invoke;
use uiautomation::actions::SelectionItem;
use uiautomation::controls::ButtonControl;
use uiautomation::controls::ListItemControl;
use windows::Win32::UI::Accessibility::UIA_ButtonControlTypeId;
use windows::Win32::UI::Accessibility::UIA_ListItemControlTypeId;

fn main() {
    let automation = UIAutomation::new().unwrap();
    let matcher = automation.create_matcher().match_name("开始").classname("Start");
    let start = matcher.find_first().unwrap();
    let button: ButtonControl = start.try_into().unwrap();
    button.invoke().unwrap();

    let matcher = automation.create_matcher().match_name("设置").classname("GridViewItem");
    let config = matcher.find_first().unwrap();
    let item: ListItemControl = config.try_into().unwrap();
    item.invoke().unwrap();

    let matcher = automation.create_matcher().match_name("设置").classname("ApplicationFrameWindow");
    let settings = matcher.find_first().unwrap();
    settings.set_focus().unwrap();
    let matcher = automation.create_matcher().from(settings.clone()).match_name("Windows 更新").control_type(UIA_ListItemControlTypeId);
    let update = matcher.find_first().unwrap();
    // println!("{}", update.get_control_type().unwrap());
    let update_item: ListItemControl = update.try_into().unwrap();
    update_item.select().unwrap();

    let matcher = automation.create_matcher().from(settings.clone()).match_name("检查更新").control_type(UIA_ButtonControlTypeId);
    let update = matcher.find_first().unwrap();
    if update.is_enabled().unwrap() {
        let button: ButtonControl = update.try_into().unwrap();
        button.invoke().unwrap();
    }
}