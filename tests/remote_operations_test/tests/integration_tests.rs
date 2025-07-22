use std::time::Duration;
use uiautomation::remote_operations::{
    RemoteOperationBuilder, RemoteOperationBatch, RemoteOperationContext
};
use uiautomation::types::{RemoteOperationMode, RemoteOperationStatus};

#[test]
fn test_remote_operations_integration() {
    // Only run this test if UI Automation is available
    if let Ok(mut context) = RemoteOperationContext::new() {
        // Test connection health check
        let is_healthy = context.check_connection_health().unwrap_or(false);
        if !is_healthy {
            println!("UI Automation connection not healthy, skipping integration test");
            return;
        }

        // Get the desktop element
        let automation = context.automation();
        if let Ok(root) = automation.get_root_element() {
            // Create a simple operation to get the desktop name
            let operation = RemoteOperationBuilder::get_property("Name")
                .with_target(root)
                .with_timeout(Duration::from_secs(5));

            // Execute it in a batch
            let batch = RemoteOperationBatch::new()
                .add_operation(operation)
                .with_mode(RemoteOperationMode::Individual);

            if let Ok(results) = batch.execute(&mut context) {
                assert_eq!(results.len(), 1);
                
                let result = &results[0];
                // The operation should at least complete (success or accessibility error)
                assert!(matches!(result.status, 
                    RemoteOperationStatus::Success | 
                    RemoteOperationStatus::AccessibilityError
                ));
                
                println!("Integration test result: {:?}", result.status);
                if let Some(ref value) = result.value {
                    if let Ok(name) = value.get_string() {
                        println!("Desktop name: '{}'", name);
                    }
                }
            }
        }

        // Test performance metrics
        let metrics = context.performance_metrics();
        assert!(metrics.total_operations > 0);
        
        println!("Total operations executed: {}", metrics.total_operations);
        println!("Success rate: {:.1}%", metrics.success_rate());
    } else {
        println!("UI Automation not available, skipping integration test");
    }
}

#[test]
fn test_batch_execution_performance() {
    if let Ok(mut context) = RemoteOperationContext::new() {
        let automation = context.automation();
        if let Ok(root) = automation.get_root_element() {
            // Test that batch mode is selected for larger operation counts
            let mut batch = RemoteOperationBatch::new()
                .with_mode(RemoteOperationMode::Auto);

            // Add 5 operations (should trigger batch mode)
            for _ in 0..5 {
                let operation = RemoteOperationBuilder::get_property("Name")
                    .with_target(root.clone())
                    .with_timeout(Duration::from_secs(2));
                batch = batch.add_operation(operation);
            }

            let start_time = std::time::Instant::now();
            if let Ok(results) = batch.execute(&mut context) {
                let execution_time = start_time.elapsed();
                
                assert_eq!(results.len(), 5);
                println!("Batch execution time: {:?}", execution_time);
                
                // Check that we saved some cross-process calls
                let metrics = context.performance_metrics();
                if metrics.cross_process_calls_saved > 0 {
                    println!("Cross-process calls saved: {}", metrics.cross_process_calls_saved);
                }
            }
        }
    }
}

#[test] 
fn test_connection_recovery() {
    if let Ok(mut context) = RemoteOperationContext::new() {
        // Test connection recovery mechanism
        let initial_health = context.check_connection_health().unwrap_or(false);
        println!("Initial connection health: {}", initial_health);
        
        // Test recovery (should succeed if automation is available)
        let recovery_result = context.recover_connection();
        match recovery_result {
            Ok(recovered) => {
                println!("Connection recovery result: {}", recovered);
                assert!(recovered);
            }
            Err(e) => {
                println!("Connection recovery failed (expected in some environments): {}", e);
            }
        }
    }
} 