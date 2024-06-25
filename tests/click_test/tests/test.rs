#[cfg(test)]
mod tests {
    use uiautomation::UIAutomation;

    #[test]
    fn test_click() {
        let automation = UIAutomation::new().unwrap();

        let matcher = automation.create_matcher().name("计算器");
        if let Ok(calc) = matcher.find_first() {
            let matcher = automation.create_matcher().name("打开导航").from_ref(&calc);
            if let Ok(navi) = matcher.find_first() {
                navi.click().unwrap();
            }
        }
    }
}