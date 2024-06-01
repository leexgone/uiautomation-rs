use uiautomation::UIAutomation;

#[tokio::main]
async fn main() {
    let handler = tokio::spawn(async {
        let automation = UIAutomation::new().unwrap();
        let root = automation.get_root_element().unwrap();

        //tokio::time::sleep(tokio::time::Duration::from_secs(1)).await;

        println!("Found it: {}", root);
    });

    let result = handler.await;

    println!("Got: {:?}", result);
}