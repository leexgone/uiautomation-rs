use std::time::Duration;
use uiautomation::processes::Process;
use uiautomation::remote_operations::{
    RemoteOperationBuilder, RemoteOperationBatch, RemoteOperationContext,
    FindCriteria
};
use uiautomation::types::RemoteOperationMode;
use uiautomation::variants::Variant;
use std::collections::HashMap;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("UI Automation Remote Operations Demo");
    println!("====================================");

    // Create a test application (Notepad) to interact with
    println!("1. Starting Notepad...");
    let _notepad = Process::create("notepad.exe")?;
    
    // Initialize UI Automation and Remote Operations context
    println!("2. Initializing Remote Operations context...");
    let mut context = RemoteOperationContext::new()?;
    
    // Get the automation instance and find the Notepad window
    let automation = context.automation();
    let root = automation.get_root_element()?;
    let matcher = automation.create_matcher().from(root).timeout(10000).classname("Notepad");
    
    if let Ok(notepad_window) = matcher.find_first() {
        println!("3. Found Notepad window: {}", notepad_window.get_name().unwrap_or_default());
        
        // Demo 1: Single Remote Operation
        println!("\n--- Demo 1: Single Remote Operation ---");
        demo_single_operation(&notepad_window, &mut context)?;
        
        // Demo 2: Batch Remote Operations
        println!("\n--- Demo 2: Batch Remote Operations ---");
        demo_batch_operations(&notepad_window, &mut context)?;
        
        // Demo 3: Performance Comparison
        println!("\n--- Demo 3: Performance Comparison ---");
        demo_performance_comparison(&notepad_window, &mut context)?;
        
        // Demo 4: Connection Health Monitoring
        println!("\n--- Demo 4: Connection Health Monitoring ---");
        demo_connection_monitoring(&mut context)?;
        
        // Demo 5: Custom Operations
        println!("\n--- Demo 5: Custom Operations ---");
        demo_custom_operations(&notepad_window, &mut context)?;
        
        // Print final performance metrics
        println!("\n--- Final Performance Metrics ---");
        print_performance_metrics(&context);
        
    } else {
        println!("Could not find Notepad window!");
    }
    
    println!("\nDemo completed. Press Enter to exit...");
    let mut input = String::new();
    std::io::stdin().read_line(&mut input)?;
    
    Ok(())
}

fn demo_single_operation(notepad_window: &uiautomation::UIElement, context: &mut RemoteOperationContext) -> Result<(), Box<dyn std::error::Error>> {
    println!("Executing single remote operation to get window name...");
    
    // Create a single operation to get the window name
    let operation = RemoteOperationBuilder::get_property("Name")
        .with_target(notepad_window.clone())
        .with_timeout(Duration::from_secs(5));
    
    // Execute it as a single-operation batch
    let batch = RemoteOperationBatch::new()
        .add_operation(operation)
        .with_mode(RemoteOperationMode::Individual);
    
    let results = batch.execute(context)?;
    
    if let Some(result) = results.first() {
        match &result.value {
            Some(value) => {
                if let Ok(name) = value.get_string() {
                    println!("✓ Got window name: '{}'", name);
                } else {
                    println!("✗ Failed to convert result to string");
                }
            }
            None => println!("✗ No value returned: {:?}", result.error_message),
        }
        println!("  Execution time: {:?}", result.execution_time);
    }
    
    Ok(())
}

fn demo_batch_operations(notepad_window: &uiautomation::UIElement, context: &mut RemoteOperationContext) -> Result<(), Box<dyn std::error::Error>> {
    println!("Executing batch of remote operations...");
    
    // Create multiple operations that would normally require separate cross-process calls
    let operations = vec![
        RemoteOperationBuilder::get_property("Name").with_target(notepad_window.clone()),
        RemoteOperationBuilder::get_property("ClassName").with_target(notepad_window.clone()),
        RemoteOperationBuilder::invoke_method("SetFocus", vec![]).with_target(notepad_window.clone()),
    ];
    
    // Execute them as a batch to save cross-process calls
    let mut batch = RemoteOperationBatch::new().with_mode(RemoteOperationMode::Batch);
    for operation in operations {
        batch = batch.add_operation(operation);
    }
    
    let start_time = std::time::Instant::now();
    let results = batch.execute(context)?;
    let batch_time = start_time.elapsed();
    
    println!("✓ Executed {} operations in batch", results.len());
    println!("  Total batch time: {:?}", batch_time);
    
    for (i, result) in results.iter().enumerate() {
        println!("  Operation {}: {:?} ({}ms)", 
                 i + 1, 
                 result.status, 
                 result.execution_time.as_millis());
        
        if let Some(ref value) = result.value {
            if let Ok(s) = value.get_string() {
                println!("    Value: '{}'", s);
            }
        }
        
        if let Some(ref error) = result.error_message {
            println!("    Error: {}", error);
        }
    }
    
    Ok(())
}

fn demo_performance_comparison(notepad_window: &uiautomation::UIElement, context: &mut RemoteOperationContext) -> Result<(), Box<dyn std::error::Error>> {
    println!("Comparing individual vs batch operation performance...");
    
    let operation_count = 5;
    
    // Test individual operations
    let start_time = std::time::Instant::now();
    for i in 0..operation_count {
        let operation = RemoteOperationBuilder::get_property("Name")
            .with_target(notepad_window.clone());
        
        let batch = RemoteOperationBatch::new()
            .add_operation(operation)
            .with_mode(RemoteOperationMode::Individual);
        
        let _ = batch.execute(context)?;
    }
    let individual_time = start_time.elapsed();
    
    // Test batch operations
    let start_time = std::time::Instant::now();
    let mut batch = RemoteOperationBatch::new().with_mode(RemoteOperationMode::Batch);
    for i in 0..operation_count {
        let operation = RemoteOperationBuilder::get_property("Name")
            .with_target(notepad_window.clone());
        batch = batch.add_operation(operation);
    }
    let _ = batch.execute(context)?;
    let batch_time = start_time.elapsed();
    
    println!("Performance comparison for {} operations:", operation_count);
    println!("  Individual mode: {:?} ({:.2} ms per operation)", 
             individual_time, 
             individual_time.as_secs_f64() * 1000.0 / operation_count as f64);
    println!("  Batch mode: {:?} ({:.2} ms per operation)", 
             batch_time, 
             batch_time.as_secs_f64() * 1000.0 / operation_count as f64);
    
    if batch_time < individual_time {
        let improvement = ((individual_time.as_secs_f64() - batch_time.as_secs_f64()) / individual_time.as_secs_f64()) * 100.0;
        println!("  ✓ Batch mode is {:.1}% faster!", improvement);
    } else {
        println!("  Individual mode performed better for this small batch");
    }
    
    Ok(())
}

fn demo_connection_monitoring(context: &mut RemoteOperationContext) -> Result<(), Box<dyn std::error::Error>> {
    println!("Demonstrating connection health monitoring...");
    
    // Check initial connection health
    let is_healthy = context.check_connection_health()?;
    println!("  Initial connection state: {:?} (healthy: {})", 
             context.connection_state(), is_healthy);
    
    // Simulate a connection issue and recovery
    println!("  Simulating connection recovery...");
    if context.recover_connection()? {
        println!("  ✓ Connection recovery successful");
        println!("  New connection state: {:?}", context.connection_state());
    } else {
        println!("  ✗ Connection recovery failed");
    }
    
    Ok(())
}

fn demo_custom_operations(notepad_window: &uiautomation::UIElement, context: &mut RemoteOperationContext) -> Result<(), Box<dyn std::error::Error>> {
    println!("Demonstrating custom operations and find criteria...");
    
    // Demo custom operation (will fail as expected)
    let custom_op = RemoteOperationBuilder::custom("CustomAction", vec![Variant::from("test_param")])
        .with_target(notepad_window.clone());
    
    let batch = RemoteOperationBatch::new().add_operation(custom_op);
    let results = batch.execute(context)?;
    
    if let Some(result) = results.first() {
        println!("  Custom operation result: {:?}", result.status);
        if let Some(ref error) = result.error_message {
            println!("  Expected error: {}", error);
        }
    }
    
    // Demo find criteria
    let mut property_conditions = HashMap::new();
    property_conditions.insert("ControlType".to_string(), Variant::from(50000i32)); // Button
    
    let find_criteria = FindCriteria {
        property_conditions,
        tree_scope: 4, // TreeScope::Descendants
        max_results: Some(10),
    };
    
    let find_op = RemoteOperationBuilder::find_elements(find_criteria)
        .with_target(notepad_window.clone());
    
    let batch = RemoteOperationBatch::new().add_operation(find_op);
    let results = batch.execute(context)?;
    
    if let Some(result) = results.first() {
        println!("  Find elements result: {:?}", result.status);
                  if let Some(ref value) = result.value {
              let count: Result<i32, _> = value.try_into();
              if let Ok(count) = count {
                  println!("  Found {} elements", count);
              }
        }
    }
    
    Ok(())
}

fn print_performance_metrics(context: &RemoteOperationContext) {
    let metrics = context.performance_metrics();
    
    println!("Total operations executed: {}", metrics.total_operations);
    println!("Successful operations: {}", metrics.successful_operations);
    println!("Failed operations: {}", metrics.failed_operations);
    println!("Success rate: {:.1}%", metrics.success_rate());
    println!("Average execution time: {:?}", metrics.average_execution_time);
    println!("Cross-process calls saved: {}", metrics.cross_process_calls_saved);
    println!("Total bytes transferred: {} bytes", metrics.total_bytes_transferred);
    
    if metrics.cross_process_calls_saved > 0 {
        println!("✓ Remote Operations provided significant performance benefits!");
    }
} 