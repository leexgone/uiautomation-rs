// Tests for remote operations functionality

#[cfg(test)]
mod tests {
    use std::time::Duration;
    use uiautomation::remote_operations::{
        RemoteOperationBuilder, RemoteOperationBatch, RemoteOperationContext,
        RemoteOperationType, FindCriteria, ConnectionState
    };
    use uiautomation::types::RemoteOperationMode;
    use uiautomation::variants::Variant;
    use std::collections::HashMap;

    #[test]
    fn test_remote_operation_builder() {
        let operation = RemoteOperationBuilder::get_property("Name")
            .with_timeout(Duration::from_secs(10));
        
        assert_eq!(operation.timeout(), Duration::from_secs(10));
        match operation.operation_type() {
            RemoteOperationType::GetProperty(prop) => assert_eq!(prop, "Name"),
            _ => panic!("Wrong operation type"),
        }
    }

    #[test]
    fn test_remote_operation_batch_creation() {
        let batch = RemoteOperationBatch::new()
            .add_operation(RemoteOperationBuilder::get_property("Name"))
            .add_operation(RemoteOperationBuilder::get_property("ClassName"))
            .with_mode(RemoteOperationMode::Batch)
            .with_timeout(Duration::from_secs(30));
        
        assert_eq!(batch.len(), 2);
        assert!(!batch.is_empty());
    }

    #[test]
    fn test_performance_metrics() {
        use uiautomation::remote_operations::PerformanceMetrics;
        
        let mut metrics = PerformanceMetrics::new();
        assert_eq!(metrics.success_rate(), 0.0);
        assert_eq!(metrics.failure_rate(), 0.0);
        
        metrics.total_operations = 10;
        metrics.successful_operations = 8;
        metrics.failed_operations = 2;
        
        assert_eq!(metrics.success_rate(), 80.0);
        assert_eq!(metrics.failure_rate(), 20.0);
        
        metrics.reset();
        assert_eq!(metrics.total_operations, 0);
        assert_eq!(metrics.successful_operations, 0);
        assert_eq!(metrics.failed_operations, 0);
    }

    #[test]
    fn test_find_criteria() {
        let mut criteria = FindCriteria {
            property_conditions: HashMap::new(),
            tree_scope: 4, // TreeScope::Descendants
            max_results: Some(10),
        };
        
        criteria.property_conditions.insert("Name".to_string(), Variant::from("Test"));
        criteria.property_conditions.insert("ControlType".to_string(), Variant::from(50000i32));
        
        assert_eq!(criteria.property_conditions.len(), 2);
        assert_eq!(criteria.max_results, Some(10));
        assert_eq!(criteria.tree_scope, 4);
    }

    #[test]
    fn test_operation_types() {
        let get_op = RemoteOperationBuilder::get_property("TestProperty");
        match get_op.operation_type() {
            RemoteOperationType::GetProperty(prop) => assert_eq!(prop, "TestProperty"),
            _ => panic!("Wrong operation type"),
        }

        let set_op = RemoteOperationBuilder::set_property("TestProperty", Variant::from("TestValue"));
        match set_op.operation_type() {
            RemoteOperationType::SetProperty(prop, _) => assert_eq!(prop, "TestProperty"),
            _ => panic!("Wrong operation type"),
        }

        let invoke_op = RemoteOperationBuilder::invoke_method("TestMethod", vec![]);
        match invoke_op.operation_type() {
            RemoteOperationType::InvokeMethod(method, _) => assert_eq!(method, "TestMethod"),
            _ => panic!("Wrong operation type"),
        }

        let custom_op = RemoteOperationBuilder::custom("CustomOp", vec![]);
        match custom_op.operation_type() {
            RemoteOperationType::Custom(name, _) => assert_eq!(name, "CustomOp"),
            _ => panic!("Wrong operation type"),
        }
    }

    #[test]
    fn test_remote_operation_context_creation() {
        // This test requires UI Automation to be available
        match RemoteOperationContext::new() {
            Ok(context) => {
                assert_eq!(*context.connection_state(), ConnectionState::Active);
                assert!(context.performance_metrics().total_operations == 0);
            }
            Err(_) => {
                // UI Automation might not be available in test environment
                println!("UI Automation not available in test environment");
            }
        }
    }

    #[test]
    fn test_batch_modes() {
        let individual_batch = RemoteOperationBatch::new()
            .with_mode(RemoteOperationMode::Individual);
        
        let batch_mode = RemoteOperationBatch::new()
            .with_mode(RemoteOperationMode::Batch);
        
        let auto_mode = RemoteOperationBatch::new()
            .with_mode(RemoteOperationMode::Auto);
        
        // Just verify the batches are created successfully
        assert_eq!(individual_batch.len(), 0);
        assert_eq!(batch_mode.len(), 0);
        assert_eq!(auto_mode.len(), 0);
    }
} 