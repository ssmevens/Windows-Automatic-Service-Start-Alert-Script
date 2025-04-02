# Wait for a reasonable amount of time after reboot (e.g., 60 seconds)
Start-Sleep -Seconds 240

# Get the initial list of services that are set to start automatically (includes both normal and delayed)
# and are not currently running.
$initialServices = Get-CimInstance -ClassName Win32_Service | Where-Object {
    $_.StartMode -eq "Auto" -and $_.State -ne "Running"
}

# Format the lists using HTML tables
function Format-ServiceTableHTML {
    param($services)
    $html = @"
<table style='border-collapse: collapse; width: 100%; font-family: Consolas, monospace; margin-bottom: 20px;'>
    <tr style='background-color: #005DAA; color: white;'>
        <th style='text-align: left; padding: 12px; border: 1px solid #003d71;'>Name</th>
        <th style='text-align: left; padding: 12px; border: 1px solid #003d71;'>DisplayName</th>
        <th style='text-align: left; padding: 12px; border: 1px solid #003d71;'>StartMode</th>
        <th style='text-align: left; padding: 12px; border: 1px solid #003d71;'>State</th>
    </tr>
"@
    
    $rowCount = 0
    foreach ($service in $services) {
        $backgroundColor = if ($rowCount % 2 -eq 0) { "#f8f9fa" } else { "#ffffff" }
        $html += @"
    <tr style='background-color: $backgroundColor;'>
        <td style='text-align: left; padding: 8px; border: 1px solid #dee2e6;'>$($service.Name)</td>
        <td style='text-align: left; padding: 8px; border: 1px solid #dee2e6;'>$($service.DisplayName)</td>
        <td style='text-align: left; padding: 8px; border: 1px solid #dee2e6;'>$($service.StartMode)</td>
        <td style='text-align: left; padding: 8px; border: 1px solid #dee2e6;'>$($service.State)</td>
    </tr>
"@
        $rowCount++
    }
    
    $html += "</table>"
    return $html
}

$initialList = Format-ServiceTableHTML $initialServices

# Initialize an array to store error messages
$errorMessages = @()

# Attempt to start each service in the initial list
foreach ($service in $initialServices) {
    try {
        Start-Service -Name $service.Name -ErrorAction Stop
    }
    catch {
        $errorMessage = "Error starting service: $($service.Name) - $_"
        # Store error message in array and also output it
        $errorMessages += $errorMessage
        Write-Output $errorMessage
    }
}

# Wait a little while to allow the services to start (e.g., 10 seconds)
Start-Sleep -Seconds 30

# Get the final list of services that are still not running (after the start attempt)
$finalServices = Get-CimInstance -ClassName Win32_Service | Where-Object {
    $_.StartMode -eq "Auto" -and $_.State -ne "Running"
}

$finalList = Format-ServiceTableHTML $finalServices

# Update email sending to use HTML
$emailBody = @"
<html>
<body style='font-family: Arial, sans-serif; max-width: 1000px; margin: 0 auto; padding: 20px;'>
    <div style='background-color: #005DAA; padding: 20px; border-radius: 5px; margin-bottom: 30px;'>
        <h1 style='color: white; margin: 0; border-bottom: 3px solid #FFD700;'>Service Status Report</h1>
    </div>

    <div style='margin-bottom: 30px;'>
        <h2 style='color: #005DAA; border-left: 4px solid #FFD700; padding-left: 10px;'>Initial list of services that were not running:</h2>
        $initialList
    </div>

    <div style='margin-bottom: 30px;'>
        <h2 style='color: #005DAA; border-left: 4px solid #FFD700; padding-left: 10px;'>List of services still not running after the start attempt:</h2>
        $finalList
    </div>

    <div style='margin-bottom: 30px;'>
        <h2 style='color: #005DAA; border-left: 4px solid #FFD700; padding-left: 10px;'>Service Start Errors:</h2>
        <pre style='background-color: #f8f9fa; padding: 15px; border-left: 4px solid #FFD700; margin: 0;'>
$($errorMessages -join "`n")
        </pre>
    </div>

    <div style='border-top: 3px solid #005DAA; padding-top: 20px; margin-top: 30px; font-size: 12px; color: #666;'>
        Generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    </div>
</body>
</html>
"@

# Configure email parameters
$smtpServer = "mail.smtp2go.com"
$from = "Service-Monitor@its-ia.com"
$port = 2525
$encryptedSMTPPass = Get-Content "C:\service-monitor\smtp-password.txt" | ConvertTo-SecureString
$smtpUser = "SRVLinux"
$smtpCred = New-Object System.Management.Automation.PSCredential($smtpUser, $encryptedSMTPPass)
$to = "alerts4tech@its-ia.com"
$hostname = $env:COMPUTERNAME
$subject = "($hostname) :  Service Start Report"

# Update the Send-MailMessage command to include -BodyAsHtml parameter
Send-MailMessage -SmtpServer $smtpServer -From $from -To $to -Subject $subject -Body $emailBody -port $port -Credential $smtpCred -BodyAsHtml
#write-host $emailBody