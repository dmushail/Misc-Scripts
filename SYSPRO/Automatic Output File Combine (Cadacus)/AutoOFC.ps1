# Output File Combine (OFC) Automation Script
# Checks for EDI files specified in DocumentFiles.txt


## Paths for document containing files to look for, log file to write to, as well as the SYSPRO exectuble.
# Replace quoted text with path to DocumentFiles.txt: 
$docFiles = Get-Content "C:\PATH\TO\DocumentFiles.txt"
# Replace quoted text with path to log file.
$log = "C:\PATH\TO\log\AutoOFC.txt"
# Replace with quoted text with path to SYSPROAuto.exe in your application directory 
$sysproExe = "C:\SYSPRO\PATH\SYSPROAuto.exe"


## SYSPRO Credentials & Parameters
# Replace the below as needed.
$operator = "OPER"
$opPass = "PASS"
$comp = "C"
$cPass = "PASS"
$prog = "EDI040"
$link = "Enterprise:NOREPORT"


## Email paremeters for error notification
# Replace recipient email address
$toEmail = "RECIPIENT@COMPANYNAME.com"
# Replace sender email address
$fromEmail = "SENDER@COMPANYNAME.com"
# Create a secure password file and replace quoted text with path to the newly created file.
$fromPass = Get-Content "SECUREPASSWORDFILE" | ConvertTo-SecureString
$authCreds = new-object `
    -typename System.Management.Automation.PSCredential `
    -argumentlist $fromEmail, $fromPass


## Start-Sleep intervals in seconds:
# How long to wait before executing OFC after detecting a file.
$waitInterval = 3
# How often to check for files.
$checkInterval = 60
# How long to wait in between tasks.
$stallInterval = 900


# Timestamp for use in log files and alert emails
function Get-TimeStamp {    
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)   
}

# Formatted date string for use in splitting log file if size limit exceeded.
function Get-DateStr {
    return '{0:yyyyMMdd}' -f (Get-Date)
}

# Send an email alert
function Send-ErrAlert {
    Send-MailMessage `
        -To $toEmail `
        -Subject "AutoOFC Failed to Run SYSPRO Output File Combine!" `
        -Body "$(Get-TimeStamp) - Failed to execute SYSPRO! Check for an EDI Report in Archived Files." `
        -UseSsl `
        -Port 587 `
        -SmtpServer 'smtp.office365.com' `
        -From $fromEmail `
        -Credential $authCreds
}


While ($true) {
    # Iterates through DocumentFiles.txt by line.
    For ($count = 0; $count -lt $docFiles.Length; $count++) {    
        $ediFile = $docFiles[$count]

        # Checks if log file exceeds a specific file size, if so, it will rename the existing log file to filename+date.
        If ($log.length -gt 5mb) {
            Rename-Item $log ($log -replace "AutoOFC", "AutoOFC $(Get-DateStr)")
        }
        
        # Checks if a file exists matching the current specified path.
        If ([System.IO.File]::Exists($ediFile)) {
            "$(Get-TimeStamp) - Outgoing file $ediFile detected." | Add-Content $log
            "$(Get-TimeStamp) - Queuing  Output File Combine for Communications Path 'Enterprise' to run in $waitInterval seconds." | Add-Content $log

            Start-Sleep -Seconds $waitInterval

            # Attempt to run SYSPRO executable
            Try {
                & $sysproExe /oper=$operator /pass=$opPass /comp=$comp /cpas=$cPass /prog=$prog /link=$link /time=60 /log

                "$(Get-TimeStamp) - Output File Combine executed successfully. Check EDI Report in Archived Files to verify." | Add-Content $log
            }
            # Write error to log if attempt results in an error, attempts to send email alert as well. 
            Catch {    
                "$(Get-TimeStamp) - FAILED to execute SYSPRO! Check for an EDI Report in Archived Files." | Add-Content $log

                # Send email alert
                Try {
                    Send-Alert

                    "$(Get-TimeStamp) - Error notification sent to $toEmail successfully!" | Add-Content $log
                    "$(Get-TimeStamp) - AutoOFC paused for $stallInterval seconds." | Add-Content $log
                }
                # Write to log if failed to send email alert
                Catch {
                    "$(Get-TimeStamp) - Failed to send email notification to $toEmail!" | Add-Content $log
                    "$(Get-TimeStamp) - AutoOFC paused for $stallInterval seconds." | Add-Content $log
                }
                
                # Pauses script for specified duration before trying again.
                "$(Get-TimeStamp) - Error. Stalling for $stallInterval seconds." | Add-Content $log
                Start-Sleep -Seconds $stallInterval
            }

        }

    }

    # Pauses script before checking for files again.
    Start-Sleep -Seconds $checkInterval
}