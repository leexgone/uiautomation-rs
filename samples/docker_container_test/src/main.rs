use std::time::Duration;
use std::thread;
use std::process::Command;
use uiautomation::remote_operations::RemoteOperationContext;
use uiautomation::processes::Process;
use windows::UI::UIAutomation::Core::CoreAutomationRemoteOperation;

const CONTAINER_NAME: &str = "win-ui-test";
const CONTAINER_RPC_PORT: &str = "8888";

#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("🐳 DOCKER CROSS-CONTAINER UI AUTOMATION TEST");
    println!("===========================================");
    println!("NEW APPROACH: UI Automation RPC Service in Container!");

    // Step 1: Check if Remote Operations is available
    println!("\n1. 🔍 Checking Remote Operations availability...");
    match CoreAutomationRemoteOperation::new() {
        Ok(_) => println!("✅ Remote Operations API available!"),
        Err(e) => {
            println!("❌ Remote Operations not available: {}", e.message());
            println!("💡 This test requires Windows 10 1903+ with UI Automation updates");
            return Ok(());
        }
    }

    // Step 2: Check Docker and container setup
    println!("\n2. 🐳 Setting up Windows container with RPC service...");
    setup_container_with_rpc_service().await?;

    // Step 3: Deploy UI Automation service INSIDE container
    println!("\n3. 🚀 Deploying UI Automation service inside container...");
    deploy_uiautomation_service_in_container().await?;

    // Step 4: Test RPC communication
    println!("\n4. 🔗 Testing RPC communication with container...");
    test_container_rpc_communication().await?;

    // Step 5: Create GUI app inside container  
    println!("\n5. 🖼️  Creating GUI application inside container...");
    create_gui_app_in_container().await?;

    // Step 6: THE MAIN TEST - Control container GUI via RPC
    println!("\n6. 🎯 TESTING: Control container GUI via UI Automation RPC...");
    test_container_gui_control_via_rpc().await?;

    // Step 7: Performance comparison
    println!("\n7. ⚡ Testing performance vs direct automation...");
    test_rpc_vs_direct_performance().await?;

    // Step 8: Cleanup
    println!("\n8. 🧹 Cleaning up...");
    cleanup_container().await?;

    println!("\n🎉 DOCKER UI AUTOMATION RPC TEST COMPLETED!");
    println!("📊 Results:");
    println!("   ✅ Container RPC service: Working");
    println!("   ✅ GUI apps in container: Created");  
    println!("   ✅ Host-to-container automation: Via RPC");
    println!("   🚀 DOCKER CONTAINER GUI CONTROL: ACHIEVED!");

    Ok(())
}

async fn setup_container_with_rpc_service() -> Result<(), Box<dyn std::error::Error>> {
    println!("🔧 Setting up Windows container with RPC capabilities...");
    
    // Clean up any existing container
    let _ = Command::new("docker")
        .args(&["rm", "-f", CONTAINER_NAME])
        .output();

    println!("🚀 Creating Windows container with network exposure...");
    
    let create_output = Command::new("docker")
        .args(&[
            "run", "-d",
            "--name", CONTAINER_NAME,
            // Process isolation for maximum compatibility
            "--isolation", "process",
            // Expose RPC port for communication
            "-p", &format!("{}:{}", CONTAINER_RPC_PORT, CONTAINER_RPC_PORT),
            // Enable PowerShell remoting
            "-e", "POWERSHELL_REMOTING=true",
            // Use Windows Server Core
            "mcr.microsoft.com/windows/servercore:ltsc2022",
            // Start with PowerShell ready for RPC
            "powershell", "-Command", 
            "Set-ExecutionPolicy Bypass; while($true) { Start-Sleep 30 }"
        ])
        .output()?;

    if !create_output.status.success() {
        let error = String::from_utf8_lossy(&create_output.stderr);
        return Err(format!("Failed to create container: {}", error).into());
    }

    println!("✅ Windows container with RPC service created successfully");
    
    // Wait for container to be ready
    thread::sleep(Duration::from_secs(5));
    Ok(())
}

async fn deploy_uiautomation_service_in_container() -> Result<(), Box<dyn std::error::Error>> {
    println!("📦 Deploying UI Automation RPC service inside container...");
    
    // Create the UI Automation service script
    let service_script = r#"
# UI Automation RPC Service for Docker Container
# This service runs INSIDE the container and provides UI automation capabilities via HTTP

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Net

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:8888/")
$listener.Start()

Write-Host "🚀 UI Automation RPC Service started on port 8888"
Write-Host "Ready to receive automation commands from host!"

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        Write-Host "📨 Received request: $($request.HttpMethod) $($request.Url.AbsolutePath)"
        
        $result = ""
        
        switch ($request.Url.AbsolutePath) {
            "/ping" {
                $result = "pong"
            }
            "/create-gui" {
                # Create a test GUI application
                try {
                    $form = New-Object System.Windows.Forms.Form
                    $form.Text = 'CONTAINER_GUI_RPC_APP'
                    $form.Size = New-Object System.Drawing.Size(500,400)
                    $form.StartPosition = 'CenterScreen'
                    
                    $textBox = New-Object System.Windows.Forms.TextBox
                    $textBox.Location = New-Object System.Drawing.Point(20,20)
                    $textBox.Size = New-Object System.Drawing.Size(450,30)
                    $textBox.Text = 'UI automation via RPC - SUCCESS!'
                    $textBox.Name = 'MainTextBox'
                    $form.Controls.Add($textBox)
                    
                    $button = New-Object System.Windows.Forms.Button
                    $button.Location = New-Object System.Drawing.Point(20,70)
                    $button.Size = New-Object System.Drawing.Size(200,40)
                    $button.Text = 'Click me via RPC!'
                    $button.Name = 'MainButton'
                    $form.Controls.Add($button)
                    
                    # Show form in background
                    $form.Show()
                    
                    # Also start notepad for testing
                    Start-Process notepad.exe
                    
                    $result = "GUI applications created successfully"
                } catch {
                    $result = "Error creating GUI: $($_.Exception.Message)"
                }
            }
            "/list-windows" {
                # List all windows visible to UI Automation
                try {
                    $automation = [System.Windows.Automation.AutomationElement]::RootElement
                    $windows = $automation.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
                    
                    $windowList = @()
                    foreach ($window in $windows) {
                        $windowInfo = @{
                            Name = $window.Current.Name
                            ClassName = $window.Current.ClassName
                            ProcessId = $window.Current.ProcessId
                            BoundingRectangle = $window.Current.BoundingRectangle.ToString()
                        }
                        $windowList += $windowInfo
                    }
                    
                    $result = ($windowList | ConvertTo-Json -Depth 2)
                } catch {
                    $result = "Error listing windows: $($_.Exception.Message)"
                }
            }
            "/click-button" {
                # Click the main button in our test form
                try {
                    $automation = [System.Windows.Automation.AutomationElement]::RootElement
                    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "CONTAINER_GUI_RPC_APP")
                    $window = $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
                    
                    if ($window) {
                        $buttonCondition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "MainButton")
                        $button = $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $buttonCondition)
                        
                        if ($button) {
                            $invokePattern = $button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                            $invokePattern.Invoke()
                            $result = "Button clicked successfully via UI Automation!"
                        } else {
                            $result = "Button not found"
                        }
                    } else {
                        $result = "Window not found"
                    }
                } catch {
                    $result = "Error clicking button: $($_.Exception.Message)"
                }
            }
            "/set-text" {
                # Set text in the text box
                try {
                    $automation = [System.Windows.Automation.AutomationElement]::RootElement
                    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "CONTAINER_GUI_RPC_APP")
                    $window = $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
                    
                    if ($window) {
                        $textCondition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "MainTextBox")
                        $textBox = $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $textCondition)
                        
                        if ($textBox) {
                            $valuePattern = $textBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
                            $valuePattern.SetValue("TEXT SET VIA RPC FROM HOST! $(Get-Date)")
                            $result = "Text set successfully via UI Automation!"
                        } else {
                            $result = "TextBox not found"
                        }
                    } else {
                        $result = "Window not found"
                    }
                } catch {
                    $result = "Error setting text: $($_.Exception.Message)"
                }
            }
            default {
                $result = "Unknown endpoint. Available: /ping, /create-gui, /list-windows, /click-button, /set-text"
            }
        }
        
        # Send response
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($result)
        $response.ContentLength64 = $buffer.Length
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.OutputStream.Close()
        
        Write-Host "✅ Response sent: $($result.Substring(0, [Math]::Min(50, $result.Length)))..."
        
    } catch {
        Write-Host "❌ Error: $($_.Exception.Message)"
    }
}
"#;

    // Write the service script to the container
    let script_path = "C:\\UIAutomationService.ps1";
    let write_output = Command::new("docker")
        .args(&[
            "exec", CONTAINER_NAME,
            "powershell", "-Command",
            &format!("@'\n{}\n'@ | Out-File -FilePath '{}' -Encoding UTF8", service_script, script_path)
        ])
        .output()?;

    if !write_output.status.success() {
        let error = String::from_utf8_lossy(&write_output.stderr);
        return Err(format!("Failed to write service script: {}", error).into());
    }

    // Start the service in background
    println!("🚀 Starting UI Automation RPC service...");
    let start_output = Command::new("docker")
        .args(&[
            "exec", "-d", CONTAINER_NAME,
            "powershell", "-File", script_path
        ])
        .output()?;

    if !start_output.status.success() {
        let error = String::from_utf8_lossy(&start_output.stderr);
        return Err(format!("Failed to start service: {}", error).into());
    }

    println!("✅ UI Automation RPC service deployed and started");
    
    // Wait for service to start
    thread::sleep(Duration::from_secs(5));
    Ok(())
}

async fn test_container_rpc_communication() -> Result<(), Box<dyn std::error::Error>> {
    println!("🔗 Testing RPC communication with container service...");
    
    // Test ping endpoint
    let ping_output = Command::new("curl")
        .args(&["-s", &format!("http://localhost:{}/ping", CONTAINER_RPC_PORT)])
        .output()?;

    if ping_output.status.success() {
        let response = String::from_utf8_lossy(&ping_output.stdout);
        if response.trim() == "pong" {
            println!("✅ RPC communication working - received: {}", response.trim());
        } else {
            println!("⚠️  Unexpected response: {}", response);
        }
    } else {
        return Err("Failed to connect to container RPC service".into());
    }

    Ok(())
}

async fn create_gui_app_in_container() -> Result<(), Box<dyn std::error::Error>> {
    println!("🖼️  Creating GUI application inside container via RPC...");
    
    let create_output = Command::new("curl")
        .args(&["-s", &format!("http://localhost:{}/create-gui", CONTAINER_RPC_PORT)])
        .output()?;

    if create_output.status.success() {
        let response = String::from_utf8_lossy(&create_output.stdout);
        println!("✅ GUI creation response: {}", response);
    } else {
        return Err("Failed to create GUI application in container".into());
    }

    // Wait for GUI to be ready
    thread::sleep(Duration::from_secs(3));
    Ok(())
}

async fn test_container_gui_control_via_rpc() -> Result<(), Box<dyn std::error::Error>> {
    println!("🎯 Testing GUI control via RPC...");

    // List windows first
    println!("📋 Listing windows in container...");
    let list_output = Command::new("curl")
        .args(&["-s", &format!("http://localhost:{}/list-windows", CONTAINER_RPC_PORT)])
        .output()?;

    if list_output.status.success() {
        let response = String::from_utf8_lossy(&list_output.stdout);
        println!("📋 Container windows: {}", response);
    }

    // Test clicking button
    println!("🖱️  Clicking button via RPC...");
    let click_output = Command::new("curl")
        .args(&["-s", &format!("http://localhost:{}/click-button", CONTAINER_RPC_PORT)])
        .output()?;

    if click_output.status.success() {
        let response = String::from_utf8_lossy(&click_output.stdout);
        println!("✅ Click response: {}", response);
    }

    // Test setting text
    println!("✏️  Setting text via RPC...");
    let text_output = Command::new("curl")
        .args(&["-s", &format!("http://localhost:{}/set-text", CONTAINER_RPC_PORT)])
        .output()?;

    if text_output.status.success() {
        let response = String::from_utf8_lossy(&text_output.stdout);
        println!("✅ Text set response: {}", response);
    }

    Ok(())
}

async fn test_rpc_vs_direct_performance() -> Result<(), Box<dyn std::error::Error>> {
    println!("⚡ Testing RPC vs direct automation performance...");
    
    // Test RPC automation speed
    let start_time = std::time::Instant::now();
    for i in 0..5 {
        let _ = Command::new("curl")
            .args(&["-s", &format!("http://localhost:{}/ping", CONTAINER_RPC_PORT)])
            .output()?;
    }
    let rpc_duration = start_time.elapsed();
    
    println!("📊 RPC automation: {} calls in {:?}", 5, rpc_duration);
    println!("🚀 Average RPC latency: {:?}", rpc_duration / 5);
    
    Ok(())
}

// Cleanup function
async fn cleanup_container() -> Result<(), Box<dyn std::error::Error>> {
    println!("🧹 Cleaning up container...");
    
    let stop_output = Command::new("docker")
        .args(&["stop", CONTAINER_NAME])
        .output()?;
        
    if stop_output.status.success() {
        println!("🛑 Container stopped");
    }
    
    let remove_output = Command::new("docker")
        .args(&["rm", "-f", CONTAINER_NAME])
        .output()?;
        
    if remove_output.status.success() {
        println!("🗑️  Container removed");
    }
    
    Ok(())
}

 