# Set the output encoding to UTF-8
$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Check if pwsh.exe is installed
$pwshPath = Get-Command pwsh -ErrorAction SilentlyContinue

if ($pwshPath) {
    # Define your PowerShell code
$scriptCode = @'

# Output an initialization message
Write-Output "Initializing ...."

# Set the output encoding to UTF-8
$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Get network statistics before the upgrade
$netStatsBeforeUpgrade = Get-NetAdapterStatistics | Where-Object { $_.ReceivedBytes -gt 0 -and $_.SentBytes -gt 0 }

# Function to prompt the user for Yes/No/Choice input with defaults
function Prompt-SelectionWithDefault {
    param (
        [string]$Message,
        [string]$DefaultOption,
        [string[]]$Options
    )

    while ($true) {
        $validOptions = $Options -join '/'
        $prompt = "$Message ($validOptions/EXIT)"
        $choice = Read-Host -Prompt $prompt

        if ([string]::IsNullOrWhiteSpace($choice)) {
            $choice = $DefaultOption
        }

        if ($Options -contains $choice) {
            return $choice
        }
        elseif ($choice -eq 'EXIT' -or $choice -eq 'Exit' -or $choice -eq 'exit') {
            exit
        }

        Write-Output "Invalid input. Please enter one of the following options: $validOptions/EXIT."
    }
}

# Function to save the error log file
function Save-ErrorLog {
    param (
        [string]$ErrorMessage
    )
    $logFilePath = "$env:USERPROFILE\Desktop\WinGet_Installation_Error_Log.txt"
    $ErrorMessage | Out-File -FilePath $logFilePath
    Write-Output "Winget is not installed. Please install it and rerun the script. Error log saved to: $logFilePath"
}

# Check if WinGet is installed
Write-Output "Checking the status of winget ...."
$wingetInstalled = Get-Command -Name winget -ErrorAction SilentlyContinue

if ($wingetInstalled) {
    Write-Output "WinGet is installed."
    Write-Output "` "
}
else {
    Write-Output "WinGet is not installed."
    # Prompt the user to choose automatic or manual installation with default set to auto
    $installChoice = Prompt-SelectionWithDefault -Message "Do you want to install WinGet automatically or manually?" -DefaultOption 'A' -Options 'A', 'M'

    if ($installChoice -eq 'A') {
        Write-Output "Installing WinGet automatically..."
        # Auto-installation code here

        try {
			# Check if Scoop is installed
			Write-Host " Checking the status of Scoop..."
			$scoopDirectory = [System.IO.Path]::Combine($env:USERPROFILE, 'scoop')

			if ((Test-Path $scoopDirectory) -and (Get-Command scoop -ErrorAction SilentlyContinue)) {
			Write-Host "Scoop is already installed."
			} else {
			Write-Host "Scoop is not installed. Installing Scoop..."
			Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
			Invoke-RestMethod -Uri https://get.scoop.sh | Invoke-Expression
			Write-Output "scoop Installed Successfully"	
			}
			Write-Output "Installing WinGet via scoop..."			
			scoop install winget
			Write-Output "WinGet installed Successfully"
        }
		
        catch {
            $errorLogMessage = @"
Failed to install WinGet. Please follow the manual installation steps:

1. Download the latest version of WinGet from the following link:
https://github.com/microsoft/winget-cli/releases

2. Locate the downloaded 'winget-cli-msixbundle.appxbundle' file.

3. Right-click the file and select 'Open'.

4. Follow the on-screen instructions to complete the installation.

If you need further assistance, please consult the WinGet documentation.
"@
			Read-Host -Prompt "Press Enter to exit"
            Save-ErrorLog -ErrorMessage $errorLogMessage
            return
        }
    }
    else {
        Write-Output "Manual installation instructions:"
        Write-Output "1. Download the latest version of WinGet from the following link:"
        Write-Output "   https://github.com/microsoft/winget-cli/releases/latest"
        Write-Output "2. Locate the downloaded file and install it according to the provided instructions."
        Write-Output "3. Once WinGet is installed, rerun this script to continue."
		Read-Host -Prompt "Press Enter to exit"
        exit
    }
}

# Add IDs to perm-skip the update for
$skipUpdate = @(

)

# Update the winget source
winget source update
Write-Output "` "

# Initialize an array to store failed or skipped app information
$failedOrSkippedApps = @()

# Define a class for software packages
class Software {
    [string]$Name
    [string]$Id
    [string]$Version
    [string]$AvailableVersion
}

# Get the current date and time
$startTime = Get-Date

# Get the available upgrades
$upgradeResult = winget upgrade -u

# Check if there are no updates
if ($upgradeResult -eq 'No installed package found matching input criteria.') {
Write-Output "No updates available. You are up-to-date."
Write-Output ""

# Keep the PowerShell window open until Enter key is pressed
Read-Host -Prompt "Press Enter to exit"
exit
}

# Initialize an array to store software package upgrades
$upgrades = @()

# Initialize variables for tracking list start and skipping irrelevant lines
$idStart = -1
$isStartList = 0

# Process each line of the upgrade result
$upgradeResult | ForEach-Object -Process {
    # Skip irrelevant lines at the beginning of the output
    if ($isStartList -lt 1 -and -not $_.StartsWith("Name") -or $_.StartsWith("---") -or $_.StartsWith("The following packages") -or $_ -match '^\s*\d+\s+upgrades\s+available\.$') {
        return
    }


    # Extract relevant information for each software package
    if ($_.StartsWith("Name")) {
        $idStart = $_.toLower().IndexOf("id")
        $isStartList = 1
        return
    }

    if ($_.Length -lt $idStart) {
        return
    }

    # Create a new instance of the Software class and assign extracted information
    $Software = [Software]::new()
    $Software.Name = $_.Substring(0, $idStart - 1)
    $info = $_.Substring($idStart) -split '\s+'
    $Software.Id = $info[0]
    $Software.Version = $info[1]
    $Software.AvailableVersion = $info[2]

    # Add the software package to the upgrades array
    $upgrades += $Software
}

# Prompt to confirm updating apps
$confirmUpdate = Prompt-SelectionWithDefault -Message "Do you want to update $($upgrades.Count) apps?" -DefaultOption 'Y' -Options 'Y', 'N', 'SHOW'

if ($confirmUpdate -eq 'SHOW') {
    # Display the list of upgrades
    $upgrades | Format-Table

    # Prompt the user to confirm if they want to update the apps
    $confirmUpdateShow = Prompt-SelectionWithDefault -Message "Do you want to update $($upgrades.Count) apps?" -DefaultOption 'Y' -Options 'Y', 'N', 'SKIP'

    # Initialize a temporary list to store apps to skip
    $tempSkipApps = @()



    while ($confirmUpdateShow -eq 'SKIP') {
        # Prompt for the ID of the app to temporarily skip
        $skipAppId = Read-Host -prompt "Enter the ID of the app to temporarily skip (Done/EXIT)"

        if ($skipAppId -eq 'DONE') {
            # Prompt to confirm skipping and updating
            $confirmUpdateShow = Prompt-SelectionWithDefault -Message "Do you want to update $($upgrades.Count) apps and temporarily skip $($tempSkipApps.Count) app(s)?" -DefaultOption 'Y' -Options 'Y', 'N'
            break
        }
        if ($skipAppId -eq '') {
            # Prompt to confirm skipping and updating
            $confirmUpdateShow = Prompt-SelectionWithDefault -Message "Do you want to update $($upgrades.Count) apps and temporarily skip $($tempSkipApps.Count) app(s)?" -DefaultOption 'Y' -Options 'Y', 'N'
            break
        }
        if ($skipAppId -eq 'exit') {
            exit
        }
		
		foreach ($id in $upgrades.Id) {
			if ($skipAppId -eq $id) {
				# Add the app ID to the temporary skip list
				Write-Output "$($skipAppId) added to temp skip list"
				$tempSkipApps += $skipAppId
			}
		}

		# Check if the app ID was not found in the upgrades list
		if ($tempSkipApps -notcontains $skipAppId) {
			Write-Output "No App ID like that"
}
}
}

# Process user's response
if ($confirmUpdate -eq 'Y' -or $confirmUpdateShow -eq 'Y') {
    # Remove temporary skip apps from the list of upgrades
    $upgrades = $upgrades | Where-Object { $tempSkipApps -notcontains $_.Id }
    # Check if the software package ID is in the skip list
    $upgrades = $upgrades | Where-Object { $skipUpdate -notcontains $_.Id }

    # Initialize arrays to store updated and skipped app information
    $updatedApps = @()

    # Get network statistics after the upgrade
    $netStatsAfterUpgrade = Get-NetAdapterStatistics | Where-Object { $_.ReceivedBytes -gt 0 -and $_.SentBytes -gt 0 }

    # Calculate the data used during the upgrade
    $dataUsedDuringUpgrade = 0
    foreach ($netStatBefore in $netStatsBeforeUpgrade) {
        $netStatAfter = $netStatsAfterUpgrade | Where-Object { $_.Name -eq $netStatBefore.Name }
        if ($null -ne $netStatAfter) {
            $dataUsedDuringUpgrade += ($netStatAfter.ReceivedBytes + $netStatAfter.SentBytes - $netStatBefore.ReceivedBytes - $netStatBefore.SentBytes) / 1MB
        }
    }

    # Initialize variables for tracking individual upgrade times and data used
    $individualUpgrades = @()

    # Get network statistics before the upgrades
    $prevNetStatsBefore = Get-NetAdapterStatistics | Where-Object { $_.ReceivedBytes -gt 0 -and $_.SentBytes -gt 0 }

    # Process each software package upgrade
    $upgrades | ForEach-Object -Process {
        Write-Output "` "
        Write-Output ("_" * 90)
        Write-Output "Going to upgrade $($_.Id)"

        # Get the start time of the individual upgrade
        $upgradeStartTime = Get-Date

        # Run the winget upgrade command
        winget upgrade --include-unknown -u --id $_.Id --silent --accept-source-agreements

        if ($LASTEXITCODE -eq 0) {
            # Successful upgrade
            $updatedApps += $_.Name

            # Get network statistics after the upgrade
            $netStatsAfter = Get-NetAdapterStatistics | Where-Object { $_.ReceivedBytes -gt 0 -and $_.SentBytes -gt 0 }

            # Calculate the data used during the upgrade
            $dataUsedPerUpgrade = 0
            foreach ($prevNetStatBefore in $prevNetStatsBefore) {
                $netStatAfter = $netStatsAfter | Where-Object { $_.Name -eq $prevNetStatBefore.Name }
                if ($null -ne $netStatAfter) {
                    $dataUsedPerUpgrade += ($netStatAfter.ReceivedBytes + $netStatAfter.SentBytes - $prevNetStatBefore.ReceivedBytes - $prevNetStatBefore.SentBytes) / 1MB
                }
            }

            # Add the individual upgrade information to the array
            $individualUpgrade = [PSCustomObject]@{
                Name     = $_.Name
                Id       = $_.Id
                DataUsed = $dataUsedPerUpgrade
                Duration = (Get-Date) - $upgradeStartTime
            }

            # Update the previous network statistics for the next upgrade
            $prevNetStatsBefore = $netStatsAfter

            # Store the individual upgrade information
            $individualUpgrades += $individualUpgrade
        }
        else {
			# Run the winget upgrade command and capture the results
			$consoleOutput = winget upgrade --include-unknown -u --id $_.Id --silent --accept-source-agreements 2>&1

			# Select the last line from the captured output
			$lastLines = $consoleOutput -split "`n" | Select-Object -Last 4	
			$lastLine = $lastLines -join '' -replace '[{}\s,\\\/]+', ' ' -replace '^-', ''	-replace ' - ', ''	
            
			# Failed upgrade
            $failedDataUsed = 0

            # Get network statistics after the upgrade
            $netStatsAfter = Get-NetAdapterStatistics | Where-Object { $_.ReceivedBytes -gt 0 -and $_.SentBytes -gt 0 }

            foreach ($prevNetStatBefore in $prevNetStatsBefore) {
                $netStatAfter = $netStatsAfter | Where-Object { $_.Name -eq $prevNetStatBefore.Name }
                if ($null -ne $netStatAfter) {
                    $failedDataUsed += ($netStatAfter.ReceivedBytes + $netStatAfter.SentBytes - $prevNetStatBefore.ReceivedBytes - $prevNetStatBefore.SentBytes) / 1MB
                }
            }

            $failedOrSkippedApps += [PSCustomObject]@{
                Name     = $_.Name
                Id       = $_.Id
                Error    = $lastLine
                DataUsed = $failedDataUsed
                Duration = (Get-Date) - $upgradeStartTime
            }

            # Update the previous network statistics for the next upgrade
            $prevNetStatsBefore = $netStatsAfter
        }
    }

    # Get the end time and calculate the total duration
    $endTime = Get-Date
    $totalDuration = $endTime - $startTime

    # Get network statistics after the upgrade
    $netStatsAfter = Get-NetAdapterStatistics | Where-Object { $_.ReceivedBytes -gt 0 -and $_.SentBytes -gt 0 }

    # Calculate the data used during the upgrade
    $dataUsed = 0
    foreach ($netStatBefore in $netStatsBeforeUpgrade) {
        $netStatAfter = $netStatsAfter | Where-Object { $_.Name -eq $netStatBefore.Name }
        if ($null -ne $netStatAfter) {
            $dataUsed += ($netStatAfter.ReceivedBytes + $netStatAfter.SentBytes - $netStatBefore.ReceivedBytes - $netStatBefore.SentBytes) / 1MB
        }
    }
	

    # Prompt to save log file
    Write-Output "` "
    $saveLog = Prompt-SelectionWithDefault -Message "Do you want to save a log file?" -DefaultOption 'Y' -Options 'Y', 'N'

    if ($saveLog -eq 'Y') {
        # Create the log file path on the desktop
        $logFilePath = "$env:USERPROFILE\Desktop\UpgradeLog.txt"

        # Check if the log file already exists
        if (Test-Path $logFilePath) {
            $overwrite = Prompt-SelectionWithDefault -Message "A log file already exists. Do you want to overwrite it?" -DefaultOption 'N' -Options 'Y', 'N'

            if ($overwrite -eq 'N') {
                $logFilePath = Join-Path $env:USERPROFILE\Desktop ("UpgradeLog" + [DateTime]::Now.ToString("yyyyMMdd_HHmmss") + ".txt")
            }
        }

        # Create the log file content
        $logContent = @"
Upgrade Log
-----------
Date: $($startTime.ToString('yyyy-MM-dd'))
Time: $($startTime.ToString('HH:mm:ss.ff')) - $($endTime.ToString('HH:mm:ss.ff'))
Total Duration: $($totalDuration.ToString("hh\:mm\:ss\.ff"))
Data Used for Winget: $($dataUsedDuringUpgrade.ToString("F5")) MB
Total Data Used: $($dataUsed.ToString("F5")) MB

Temporarily Skipped apps:
$($tempSkipApps -join "`r`n")

Updated Apps:
$($individualUpgrades | Select-Object @{Name='Name'; Expression={$_.Name + "  "}},  @{Name='Id'; Expression={$_.Id + "  "}}, @{Name='DataUsed'; Expression={$_.DataUsed.ToString("F5") + " MB  "}}, @{Name='Duration'; Expression={$_.Duration.ToString("hh\:mm\:ss\.ff") + "  "}}| Format-Table -AutoSize | Out-String -Width ([int]::MaxValue))

Update Failed:
$($failedOrSkippedApps | Select-Object @{Name='Name'; Expression={$_.Name + "  "}},  @{Name='Id'; Expression={$_.Id + "  "}}, @{Name='DataUsed'; Expression={$_.DataUsed.ToString("F5") + " MB  "}}, @{Name='Duration'; Expression={$_.Duration.ToString("hh\:mm\:ss\.ff") + "  "}}, @{Name='Error'; Expression={$_.Error -replace '{(.*?)}', '$1'}}| Format-Table -AutoSize | Out-String -Width ([int]::MaxValue))

"@
        # Write the log content to the log file
        $logContent | Out-File -FilePath $logFilePath

        # Display the log file path
        Write-Output "Upgrade log file saved at: $logFilePath"
    }
}

# Keep the PowerShell window open until Enter key is pressed
Read-Host -Prompt "Press Enter to exit"
exit

'@

# Create a temporary script file
$scriptPath = [System.IO.Path]::Combine($env:TEMP, "TemporaryScript.ps1")
$scriptCode | Out-File -FilePath $scriptPath -Encoding UTF8

# Open Windows Terminal and execute the script
Start-Process "wt" -ArgumentList "pwsh.exe -File `"$scriptPath`""

} else {
    Write-Host "PowerShell Core (pwsh.exe) is not installed."

    # Display installation instructions
    Write-Host "To install PowerShell Core, follow these steps:"
    Write-Host "1. Visit the official PowerShell GitHub releases page: https://github.com/PowerShell/PowerShell/releases"
    Write-Host "2. Download the latest stable release for your operating system (Windows, macOS, or Linux)."
    Write-Host "3. Follow the installation instructions provided for your specific platform."

   }

