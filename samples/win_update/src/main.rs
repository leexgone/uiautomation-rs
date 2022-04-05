use uiautomation::UIAutomation;
use uiautomation::controls::ButtonControl;
use uiautomation::controls::ListItemControl;

fn main() {
    let automation = UIAutomation::new().unwrap();
    let matcher = automation.create_matcher().match_name("开始").match_classname("Start");
    let start = matcher.find_first().unwrap();
    let button: ButtonControl = start.try_into().unwrap();
    button.click().unwrap();

    let matcher = automation.create_matcher().match_name("设置").match_classname("GridViewItem");
    let config = matcher.find_first().unwrap();
    let item: ListItemControl = config.try_into().unwrap();
    item.click().unwrap();
}