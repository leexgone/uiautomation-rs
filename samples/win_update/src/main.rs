#![windows_subsystem = "windows"]

use uiautomation::Result;
use uiautomation::UIAutomation;
use uiautomation::UIElement;
use uiautomation::actions::Invoke;
use uiautomation::actions::SelectionItem;
use uiautomation::actions::Toggle;
use uiautomation::controls::ButtonControl;
use uiautomation::controls::Control;
use uiautomation::controls::ListItemControl;
use uiautomation::controls::PaneControl;
use uiautomation::controls::ToolBarControl;
use uiautomation::controls::WindowControl;
use uiautomation::dialogs::show_error;
use uiautomation::filters::NameFilter;
use uiautomation::filters::OrFilter;

fn main() {
    let ret = auto_update();
    if let Err(ref e) = ret {
        show_error(if e.code() == 0 { "遇到未知的错误" } else { e.message() }, "win-update");

        ret.unwrap();
    }
}

fn auto_update() -> Result<()> {
    let automation = UIAutomation::new()?;

    if let Ok(start) = automation.create_matcher().match_name("开始").classname("ToggleButton").find_first() {     // 尝试Win11
        let button: ButtonControl = start.try_into()?;
        button.toggle()?;
    } else if let Ok(start) = automation.create_matcher().match_name("开始").classname("Start").find_first() {      // 尝试Win10
        let button: ButtonControl = start.try_into()?;
        button.invoke()?;
    } else {
        return Err("无法定位开始按钮".into());
    }

    let matcher = automation.create_matcher().match_name("设置").classname("GridViewItem");
    let config = matcher.find_first()?;
    let item: ListItemControl = config.try_into()?;
    item.invoke()?;

    let matcher = automation.create_matcher().match_name("设置").classname("ApplicationFrameWindow");
    let settings = matcher.find_first()?;
    
    let window: WindowControl = settings.clone().try_into()?;
    window.set_foregrand()?;
    settings.set_focus()?;

    let matcher = automation.create_matcher().from(settings.clone()).contains_name("Windows 更新").control_type(ListItemControl::TYPE).timeout(10000);
    let update = matcher.find_first()?;
    // println!("{}", update.get_control_type()?);
    let update_item: ListItemControl = update.try_into()?;
    update_item.select()?;

    let filter = OrFilter {
        left: Box::new(NameFilter { value: String::from("检查更新"), casesensitive: false, partial: true }),
        right: Box::new(NameFilter { value: String::from("下载并安装"), casesensitive: false, partial: true }),
    };
    let matcher = automation.create_matcher().from(settings.clone()).timeout(5000) //.debug(true)
        .filter(Box::new(filter))
        // .match_name("检查更新")
        .control_type(ButtonControl::TYPE)
        .filter_fn(Box::new(|e: &UIElement| {
            e.is_enabled()
        }));
    let update = matcher.find_first()?;
    if update.is_enabled()? {
        let button: ButtonControl = update.try_into()?;
        button.invoke()?;
    } else {
        return Err("调用更新失败！".into())
    }

    if let Ok(taskbar) = automation.create_matcher().name("运行中的应用程序").control_type(ToolBarControl::TYPE).find_first() { // Win10
        let matcher = automation.create_matcher().from(taskbar).contains_name("设置").control_type(ButtonControl::TYPE);
        if let Ok(settings_button) = matcher.find_first() {
            settings_button.click()?;
        }
    } else if let Ok(taskbar) = automation.create_matcher().name("任务栏").control_type(PaneControl::TYPE).find_first() {       // Win11
        let matcher = automation.create_matcher().from(taskbar).contains_name("设置").control_type(ButtonControl::TYPE);
        if let Ok(settings_button) = matcher.find_first() {
            settings_button.click()?;
        }
    };

    Ok(())
}