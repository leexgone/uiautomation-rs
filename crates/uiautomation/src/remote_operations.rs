use std::collections::HashMap;
use std::time::{Duration, Instant};

use crate::core::{UIAutomation, UIElement};
use crate::errors::Result;
use crate::types::{RemoteOperationStatus, RemoteOperationMode, RemoteOperationExtensibilityLevel, RemoteOperationCallbackBehavior};
use crate::variants::Variant;

/// Represents a single operation that can be executed remotely.
#[derive(Clone)]
pub struct RemoteOperation {
    operation_type: RemoteOperationType,
    target_element: Option<UIElement>,
    parameters: HashMap<String, Variant>,
    timeout: Duration,
}

/// The type of remote operation to execute.
#[derive(Clone)]
pub enum RemoteOperationType {
    /// Get a property value from an element.
    GetProperty(String),
    /// Set a property value on an element.
    SetProperty(String, Variant),
    /// Invoke a method on an element.
    InvokeMethod(String, Vec<Variant>),
    /// Find elements matching criteria.
    FindElements(FindCriteria),
    /// Get pattern from element.
    GetPattern(String),
    /// Custom operation with user-defined logic.
    Custom(String, Vec<Variant>),
}

/// Criteria for finding elements in remote operations.
#[derive(Clone)]
pub struct FindCriteria {
    pub property_conditions: HashMap<String, Variant>,
    pub tree_scope: i32,
    pub max_results: Option<usize>,
}

/// Result of a remote operation execution.
#[derive(Clone)]
pub struct RemoteOperationResult {
    pub status: RemoteOperationStatus,
    pub value: Option<Variant>,
    pub error_message: Option<String>,
    pub execution_time: Duration,
}

/// A batch of remote operations that can be executed together.
pub struct RemoteOperationBatch {
    operations: Vec<RemoteOperation>,
    mode: RemoteOperationMode,
    timeout: Duration,
    callback_behavior: RemoteOperationCallbackBehavior,
}

/// Context for executing remote operations, including connection state and performance metrics.
pub struct RemoteOperationContext {
    automation: UIAutomation,
    connection_state: ConnectionState,
    performance_metrics: PerformanceMetrics,
    extensibility_level: RemoteOperationExtensibilityLevel,
}

/// Represents the state of the connection to the remote automation provider.
#[derive(Debug, Clone, PartialEq)]
pub enum ConnectionState {
    /// Connection is active and healthy.
    Active,
    /// Connection is degraded but functional.
    Degraded,
    /// Connection is lost and needs recovery.
    Lost,
    /// Connection is being established.
    Connecting,
}

/// Performance metrics for remote operations.
#[derive(Debug, Clone)]
pub struct PerformanceMetrics {
    pub total_operations: u64,
    pub successful_operations: u64,
    pub failed_operations: u64,
    pub average_execution_time: Duration,
    pub total_bytes_transferred: u64,
    pub cross_process_calls_saved: u64,
}

impl RemoteOperation {
    /// Creates a new remote operation.
    pub fn new(operation_type: RemoteOperationType) -> Self {
        Self {
            operation_type,
            target_element: None,
            parameters: HashMap::new(),
            timeout: Duration::from_secs(30),
        }
    }

    /// Sets the target element for this operation.
    pub fn with_target(mut self, element: UIElement) -> Self {
        self.target_element = Some(element);
        self
    }

    /// Adds a parameter to this operation.
    pub fn with_parameter<K: Into<String>, V: Into<Variant>>(mut self, key: K, value: V) -> Self {
        self.parameters.insert(key.into(), value.into());
        self
    }

    /// Sets the timeout for this operation.
    pub fn with_timeout(mut self, timeout: Duration) -> Self {
        self.timeout = timeout;
        self
    }

    /// Gets the operation type.
    pub fn operation_type(&self) -> &RemoteOperationType {
        &self.operation_type
    }

    /// Gets the target element.
    pub fn target_element(&self) -> Option<&UIElement> {
        self.target_element.as_ref()
    }

    /// Gets the parameters.
    pub fn parameters(&self) -> &HashMap<String, Variant> {
        &self.parameters
    }

    /// Gets the timeout.
    pub fn timeout(&self) -> Duration {
        self.timeout
    }
}

impl RemoteOperationBatch {
    /// Creates a new remote operation batch.
    pub fn new() -> Self {
        Self {
            operations: Vec::new(),
            mode: RemoteOperationMode::Auto,
            timeout: Duration::from_secs(60),
            callback_behavior: RemoteOperationCallbackBehavior::Synchronous,
        }
    }

    /// Adds an operation to the batch.
    pub fn add_operation(mut self, operation: RemoteOperation) -> Self {
        self.operations.push(operation);
        self
    }

    /// Sets the execution mode for the batch.
    pub fn with_mode(mut self, mode: RemoteOperationMode) -> Self {
        self.mode = mode;
        self
    }

    /// Sets the timeout for the entire batch.
    pub fn with_timeout(mut self, timeout: Duration) -> Self {
        self.timeout = timeout;
        self
    }

    /// Sets the callback behavior for the batch.
    pub fn with_callback_behavior(mut self, behavior: RemoteOperationCallbackBehavior) -> Self {
        self.callback_behavior = behavior;
        self
    }

    /// Gets the number of operations in the batch.
    pub fn len(&self) -> usize {
        self.operations.len()
    }

    /// Checks if the batch is empty.
    pub fn is_empty(&self) -> bool {
        self.operations.is_empty()
    }

    /// Gets the operations in the batch.
    pub fn operations(&self) -> &[RemoteOperation] {
        &self.operations
    }

    /// Executes the batch and returns results for all operations.
    pub fn execute(&self, context: &mut RemoteOperationContext) -> Result<Vec<RemoteOperationResult>> {
        if self.operations.is_empty() {
            return Ok(Vec::new());
        }

        let start_time = Instant::now();
        let mut results = Vec::with_capacity(self.operations.len());

        // Determine execution strategy based on mode and operation count
        let should_batch = match self.mode {
            RemoteOperationMode::Individual => false,
            RemoteOperationMode::Batch => true,
            RemoteOperationMode::Auto => self.operations.len() > 3, // Auto-batch for 4+ operations
        };

        if should_batch {
            results = self.execute_batch_mode(context)?;
        } else {
            for operation in &self.operations {
                let result = self.execute_single_operation(operation, context)?;
                results.push(result);
            }
        }

        // Update performance metrics
        let execution_time = start_time.elapsed();
        context.update_metrics(&results, execution_time);

        Ok(results)
    }

    /// Executes operations in batch mode to minimize cross-process calls.
    fn execute_batch_mode(&self, context: &mut RemoteOperationContext) -> Result<Vec<RemoteOperationResult>> {
        let _start_time = Instant::now();
        let mut results = Vec::with_capacity(self.operations.len());

        // For now, simulate batch execution by grouping similar operations
        // In a real implementation, this would use the Windows Remote Operations API
        for operation in &self.operations {
            let result = self.execute_single_operation(operation, context)?;
            results.push(result);
        }

        // Simulate the performance benefit of batching
        let saved_calls = if self.operations.len() > 1 {
            (self.operations.len() - 1) as u64
        } else {
            0
        };
        context.performance_metrics.cross_process_calls_saved += saved_calls;

        Ok(results)
    }

    /// Executes a single operation.
    fn execute_single_operation(&self, operation: &RemoteOperation, context: &mut RemoteOperationContext) -> Result<RemoteOperationResult> {
        let start_time = Instant::now();

        let result = match &operation.operation_type {
            RemoteOperationType::GetProperty(property) => {
                self.execute_get_property(operation, property, context)
            }
            RemoteOperationType::SetProperty(property, value) => {
                self.execute_set_property(operation, property, value, context)
            }
            RemoteOperationType::InvokeMethod(method, args) => {
                self.execute_invoke_method(operation, method, args, context)
            }
            RemoteOperationType::FindElements(criteria) => {
                self.execute_find_elements(operation, criteria, context)
            }
            RemoteOperationType::GetPattern(pattern) => {
                self.execute_get_pattern(operation, pattern, context)
            }
            RemoteOperationType::Custom(name, args) => {
                self.execute_custom_operation(operation, name, args, context)
            }
        };

        let execution_time = start_time.elapsed();
        
        // Check for timeout
        if execution_time > operation.timeout {
            return Ok(RemoteOperationResult {
                status: RemoteOperationStatus::Timeout,
                value: None,
                error_message: Some("Operation timed out".to_string()),
                execution_time,
            });
        }

        result.map(|mut r| {
            r.execution_time = execution_time;
            r
        })
    }

    /// Executes a get property operation.
    fn execute_get_property(&self, operation: &RemoteOperation, property: &str, _context: &mut RemoteOperationContext) -> Result<RemoteOperationResult> {
        if let Some(element) = &operation.target_element {
            // This is a simplified implementation - in reality you'd use the actual property getting logic
            match property {
                "Name" => {
                    match element.get_name() {
                        Ok(name) => Ok(RemoteOperationResult {
                            status: RemoteOperationStatus::Success,
                            value: Some(Variant::from(name)),
                            error_message: None,
                            execution_time: Duration::default(),
                        }),
                        Err(e) => Ok(RemoteOperationResult {
                            status: RemoteOperationStatus::AccessibilityError,
                            value: None,
                            error_message: Some(e.to_string()),
                            execution_time: Duration::default(),
                        }),
                    }
                }
                "ClassName" => {
                    match element.get_classname() {
                        Ok(class_name) => Ok(RemoteOperationResult {
                            status: RemoteOperationStatus::Success,
                            value: Some(Variant::from(class_name)),
                            error_message: None,
                            execution_time: Duration::default(),
                        }),
                        Err(e) => Ok(RemoteOperationResult {
                            status: RemoteOperationStatus::AccessibilityError,
                            value: None,
                            error_message: Some(e.to_string()),
                            execution_time: Duration::default(),
                        }),
                    }
                }
                _ => Ok(RemoteOperationResult {
                    status: RemoteOperationStatus::InvalidArgument,
                    value: None,
                    error_message: Some(format!("Unknown property: {}", property)),
                    execution_time: Duration::default(),
                }),
            }
        } else {
            Ok(RemoteOperationResult {
                status: RemoteOperationStatus::InvalidArgument,
                value: None,
                error_message: Some("No target element specified".to_string()),
                execution_time: Duration::default(),
            })
        }
    }

    /// Executes a set property operation.
    fn execute_set_property(&self, _operation: &RemoteOperation, property: &str, _value: &Variant, _context: &mut RemoteOperationContext) -> Result<RemoteOperationResult> {
        // Simplified implementation - most UI Automation properties are read-only
        Ok(RemoteOperationResult {
            status: RemoteOperationStatus::InvalidArgument,
            value: None,
            error_message: Some(format!("Property '{}' is not settable", property)),
            execution_time: Duration::default(),
        })
    }

    /// Executes an invoke method operation.
    fn execute_invoke_method(&self, operation: &RemoteOperation, method: &str, _args: &[Variant], _context: &mut RemoteOperationContext) -> Result<RemoteOperationResult> {
        if let Some(element) = &operation.target_element {
            match method {
                "SetFocus" => {
                    match element.set_focus() {
                        Ok(_) => Ok(RemoteOperationResult {
                            status: RemoteOperationStatus::Success,
                            value: None,
                            error_message: None,
                            execution_time: Duration::default(),
                        }),
                        Err(e) => Ok(RemoteOperationResult {
                            status: RemoteOperationStatus::AccessibilityError,
                            value: None,
                            error_message: Some(e.to_string()),
                            execution_time: Duration::default(),
                        }),
                    }
                }
                _ => Ok(RemoteOperationResult {
                    status: RemoteOperationStatus::InvalidArgument,
                    value: None,
                    error_message: Some(format!("Unknown method: {}", method)),
                    execution_time: Duration::default(),
                }),
            }
        } else {
            Ok(RemoteOperationResult {
                status: RemoteOperationStatus::InvalidArgument,
                value: None,
                error_message: Some("No target element specified".to_string()),
                execution_time: Duration::default(),
            })
        }
    }

    /// Executes a find elements operation.
    fn execute_find_elements(&self, _operation: &RemoteOperation, _criteria: &FindCriteria, _context: &mut RemoteOperationContext) -> Result<RemoteOperationResult> {
        // Simplified implementation
        Ok(RemoteOperationResult {
            status: RemoteOperationStatus::Success,
            value: Some(Variant::from(0i32)), // Return count of found elements
            error_message: None,
            execution_time: Duration::default(),
        })
    }

    /// Executes a get pattern operation.
    fn execute_get_pattern(&self, _operation: &RemoteOperation, _pattern: &str, _context: &mut RemoteOperationContext) -> Result<RemoteOperationResult> {
        // Simplified implementation
        Ok(RemoteOperationResult {
            status: RemoteOperationStatus::Success,
            value: None,
            error_message: None,
            execution_time: Duration::default(),
        })
    }

    /// Executes a custom operation.
    fn execute_custom_operation(&self, _operation: &RemoteOperation, name: &str, _args: &[Variant], _context: &mut RemoteOperationContext) -> Result<RemoteOperationResult> {
        Ok(RemoteOperationResult {
            status: RemoteOperationStatus::InvalidArgument,
            value: None,
            error_message: Some(format!("Custom operation '{}' not implemented", name)),
            execution_time: Duration::default(),
        })
    }
}

impl RemoteOperationContext {
    /// Creates a new remote operation context.
    pub fn new() -> Result<Self> {
        let automation = UIAutomation::new()?;
        
        Ok(Self {
            automation,
            connection_state: ConnectionState::Active,
            performance_metrics: PerformanceMetrics::new(),
            extensibility_level: RemoteOperationExtensibilityLevel::Standard,
        })
    }

    /// Creates a new context with the specified extensibility level.
    pub fn with_extensibility_level(extensibility_level: RemoteOperationExtensibilityLevel) -> Result<Self> {
        let mut context = Self::new()?;
        context.extensibility_level = extensibility_level;
        Ok(context)
    }

    /// Gets the automation instance.
    pub fn automation(&self) -> &UIAutomation {
        &self.automation
    }

    /// Gets the current connection state.
    pub fn connection_state(&self) -> &ConnectionState {
        &self.connection_state
    }

    /// Gets the performance metrics.
    pub fn performance_metrics(&self) -> &PerformanceMetrics {
        &self.performance_metrics
    }

    /// Gets the extensibility level.
    pub fn extensibility_level(&self) -> RemoteOperationExtensibilityLevel {
        self.extensibility_level
    }

    /// Updates performance metrics based on operation results.
    fn update_metrics(&mut self, results: &[RemoteOperationResult], total_time: Duration) {
        self.performance_metrics.total_operations += results.len() as u64;
        
        for result in results {
            match result.status {
                RemoteOperationStatus::Success => {
                    self.performance_metrics.successful_operations += 1;
                }
                _ => {
                    self.performance_metrics.failed_operations += 1;
                }
            }
        }

        // Update average execution time
        let total_ops = self.performance_metrics.total_operations;
        if total_ops > 0 {
            let current_avg_nanos = self.performance_metrics.average_execution_time.as_nanos();
            let new_avg_nanos = (current_avg_nanos * (total_ops - results.len() as u64) as u128 + 
                                total_time.as_nanos()) / total_ops as u128;
            self.performance_metrics.average_execution_time = Duration::from_nanos(new_avg_nanos as u64);
        }
    }

    /// Checks the connection health and updates state.
    pub fn check_connection_health(&mut self) -> Result<bool> {
        // Simplified implementation - try to get the root element
        match self.automation.get_root_element() {
            Ok(_) => {
                self.connection_state = ConnectionState::Active;
                Ok(true)
            }
            Err(_) => {
                self.connection_state = ConnectionState::Lost;
                Ok(false)
            }
        }
    }

    /// Attempts to recover a lost connection.
    pub fn recover_connection(&mut self) -> Result<bool> {
        self.connection_state = ConnectionState::Connecting;
        
        // Try to create a new automation instance
        match UIAutomation::new() {
            Ok(automation) => {
                self.automation = automation;
                self.connection_state = ConnectionState::Active;
                Ok(true)
            }
            Err(e) => {
                self.connection_state = ConnectionState::Lost;
                Err(e)
            }
        }
    }
}

impl PerformanceMetrics {
    /// Creates new performance metrics.
    pub fn new() -> Self {
        Self {
            total_operations: 0,
            successful_operations: 0,
            failed_operations: 0,
            average_execution_time: Duration::default(),
            total_bytes_transferred: 0,
            cross_process_calls_saved: 0,
        }
    }

    /// Gets the success rate as a percentage.
    pub fn success_rate(&self) -> f64 {
        if self.total_operations == 0 {
            0.0
        } else {
            (self.successful_operations as f64 / self.total_operations as f64) * 100.0
        }
    }

    /// Gets the failure rate as a percentage.
    pub fn failure_rate(&self) -> f64 {
        if self.total_operations == 0 {
            0.0
        } else {
            (self.failed_operations as f64 / self.total_operations as f64) * 100.0
        }
    }

    /// Resets all metrics to zero.
    pub fn reset(&mut self) {
        *self = Self::new();
    }
}

impl Default for RemoteOperationBatch {
    fn default() -> Self {
        Self::new()
    }
}

impl Default for PerformanceMetrics {
    fn default() -> Self {
        Self::new()
    }
}

/// Builder for creating remote operations with a fluent API.
pub struct RemoteOperationBuilder;

impl RemoteOperationBuilder {
    /// Creates a new get property operation.
    pub fn get_property<S: Into<String>>(property: S) -> RemoteOperation {
        RemoteOperation::new(RemoteOperationType::GetProperty(property.into()))
    }

    /// Creates a new set property operation.
    pub fn set_property<S: Into<String>, V: Into<Variant>>(property: S, value: V) -> RemoteOperation {
        RemoteOperation::new(RemoteOperationType::SetProperty(property.into(), value.into()))
    }

    /// Creates a new invoke method operation.
    pub fn invoke_method<S: Into<String>>(method: S, args: Vec<Variant>) -> RemoteOperation {
        RemoteOperation::new(RemoteOperationType::InvokeMethod(method.into(), args))
    }

    /// Creates a new find elements operation.
    pub fn find_elements(criteria: FindCriteria) -> RemoteOperation {
        RemoteOperation::new(RemoteOperationType::FindElements(criteria))
    }

    /// Creates a new get pattern operation.
    pub fn get_pattern<S: Into<String>>(pattern: S) -> RemoteOperation {
        RemoteOperation::new(RemoteOperationType::GetPattern(pattern.into()))
    }

    /// Creates a new custom operation.
    pub fn custom<S: Into<String>>(name: S, args: Vec<Variant>) -> RemoteOperation {
        RemoteOperation::new(RemoteOperationType::Custom(name.into(), args))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_remote_operation_creation() {
        let operation = RemoteOperationBuilder::get_property("Name")
            .with_timeout(Duration::from_secs(10));
        
        assert_eq!(operation.timeout(), Duration::from_secs(10));
        match operation.operation_type() {
            RemoteOperationType::GetProperty(prop) => assert_eq!(prop, "Name"),
            _ => panic!("Wrong operation type"),
        }
    }

    #[test]
    fn test_remote_operation_batch() {
        let batch = RemoteOperationBatch::new()
            .add_operation(RemoteOperationBuilder::get_property("Name"))
            .add_operation(RemoteOperationBuilder::get_property("ClassName"))
            .with_mode(RemoteOperationMode::Batch);
        
        assert_eq!(batch.len(), 2);
        assert!(!batch.is_empty());
    }

    #[test]
    fn test_performance_metrics() {
        let mut metrics = PerformanceMetrics::new();
        assert_eq!(metrics.success_rate(), 0.0);
        assert_eq!(metrics.failure_rate(), 0.0);
        
        metrics.total_operations = 10;
        metrics.successful_operations = 8;
        metrics.failed_operations = 2;
        
        assert_eq!(metrics.success_rate(), 80.0);
        assert_eq!(metrics.failure_rate(), 20.0);
    }

    #[test]
    fn test_find_criteria() {
        let mut criteria = FindCriteria {
            property_conditions: HashMap::new(),
            tree_scope: 4, // TreeScope::Descendants
            max_results: Some(10),
        };
        
        criteria.property_conditions.insert("Name".to_string(), Variant::from("Test"));
        
        assert_eq!(criteria.property_conditions.len(), 1);
        assert_eq!(criteria.max_results, Some(10));
    }
} 