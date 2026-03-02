<#
    Ollama Windows Service Installer using NSSM
    -------------------------------------------
    This script uses NSSM (Non-Sucking Service Manager) to wrap ollama.exe as a Windows service.
    NSSM will be automatically downloaded if not present.

    Run as Administrator:
    Start-Process powershell -Verb RunAs -ArgumentList "-NoExit", "-Command", "cd 'C:\myCode\gitHub\mcp-ollama-python\scripts'; .\win_ollama_service_install_nssm.ps1"

    Features:
    - Auto-downloads NSSM if needed
    - Detects Ollama installation path
    - Supports custom port (default: 11434)
    - Supports custom model storage path
    - Creates firewall rule
    - Supports uninstall mode
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidatePattern('^[a-zA-Z0-9 _-]+$')]
    [ValidateLength(1, 80)]
    [string]$ServiceName = "ollama",

    [Parameter(Mandatory=$false)]
    [ValidateLength(1, 256)]
    [string]$DisplayName = "Ollama Service",

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 65535)]
    [int]$Port = 11434,

    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if ($_ -and -not (Test-Path $_ -PathType Container)) {
            throw "Model path does not exist or is not a directory"
        }
        return $true
    })]
    [string]$ModelPath,

    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if ($_ -and -not (Test-Path $_ -PathType Leaf)) {
            throw "Ollama path does not exist or is not a file"
        }
        return $true
    })]
    [string]$OllamaPath,

    [Parameter(Mandatory=$false)]
    [string]$NssmPath,

    [Parameter(Mandatory=$false)]
    [switch]$Install,

    [Parameter(Mandatory=$false)]
    [switch]$Uninstall,

    [Parameter(Mandatory=$false)]
    [switch]$Status,

    [Parameter(Mandatory=$false)]
    [switch]$Stop,

    [Parameter(Mandatory=$false)]
    [switch]$Start,

    [Parameter(Mandatory=$false)]
    [switch]$Silent
)

# Stable per-machine data directory for logs and install log
$DataDir = "$env:ProgramData\Ollama"
if (-not (Test-Path $DataDir)) {
    New-Item -Path $DataDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
}

$LogFile = "$DataDir\OllamaServiceInstall.log"

# NSSM path: -NssmPath param → C:\Program Files\NSSM\nssm.exe → script dir (dev fallback)
if (-not $NssmPath) {
    $NssmPath = "$env:ProgramFiles\NSSM\nssm.exe"
    $localNssm = Join-Path $PSScriptRoot "nssm.exe"
    if (-not (Test-Path $NssmPath) -and (Test-Path $localNssm)) {
        $NssmPath = $localNssm
    }
}

function Log {
    param(
        [string]$Message,
        [switch]$NoConsole,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "$timestamp [$Level] $Message"

    try {
        $line | Out-File -FilePath $LogFile -Append -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        # Fallback to event log if file logging fails
        try {
            Write-EventLog -LogName Application -Source "OllamaInstaller" -EventId 1001 -EntryType Warning -Message "Failed to write to log file: $Message" -ErrorAction SilentlyContinue
        }
        catch {
            # Last resort - write to console even in silent mode
            Write-Host "LOG ERROR: $Message" -ForegroundColor Red
        }
    }

    if (-not $Silent -and -not $NoConsole) {
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow }
            'DEBUG' { Write-Host $Message -ForegroundColor Gray }
            default { Write-Host $Message }
        }
    }
}

function Show-Help {
    $scriptName = $PSCommandPath | Split-Path -Leaf
    Write-Host ""
    Write-Host "Ollama Windows Service Manager (NSSM)" -ForegroundColor Cyan
    Write-Host "======================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "USAGE:" -ForegroundColor Yellow
    Write-Host "  .$scriptName -Install   [options]   Install and start the service (requires Admin)"
    Write-Host "  .$scriptName -Uninstall [options]   Stop and remove the service  (requires Admin)"
    Write-Host "  .$scriptName -Start                 Start the service             (requires Admin)"
    Write-Host "  .$scriptName -Stop                  Stop the service              (requires Admin)"
    Write-Host "  .$scriptName -Status                Show current service status"
    Write-Host ""
    Write-Host "OPTIONS (used with -Install):" -ForegroundColor Yellow
    Write-Host "  -ServiceName  <name>   Service name (default: ollama)"
    Write-Host "  -DisplayName  <name>   Display name (default: 'Ollama Service')"
    Write-Host "  -Port         <int>    Listening port (default: 11434)"
    Write-Host "  -ModelPath    <path>   Custom model storage directory"
    Write-Host "  -OllamaPath   <path>   Path to ollama.exe (auto-detected if omitted)"
    Write-Host "  -NssmPath     <path>   Path to nssm.exe (default: '$env:ProgramFiles\NSSM\nssm.exe')"
    Write-Host "  -Silent                Suppress console output"
    Write-Host ""
    Write-Host "EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  .$scriptName -Install"
    Write-Host "  .$scriptName -Install -Port 12345 -ModelPath 'D:\models'"
    Write-Host "  .$scriptName -Install -ServiceName myollama"
    Write-Host "  .$scriptName -Status"
    Write-Host "  .$scriptName -Stop"
    Write-Host "  .$scriptName -Uninstall"
    Write-Host ""
}

# --- Guard: catch action words mistakenly passed as service name ---
$reservedWords = @('start', 'stop', 'status', 'uninstall', 'install', 'restart')
if ($reservedWords -contains $ServiceName.ToLower()) {
    Write-Host "ERROR: '$ServiceName' is not a valid service name - did you mean: -$ServiceName ?" -ForegroundColor Red
    Show-Help
    exit 1
}

Log "=== Ollama Service Installer (NSSM) started ==="
Log "Parameters: ServiceName=$ServiceName, Port=$Port, ModelPath=$ModelPath, OllamaPath=$OllamaPath, Uninstall=$Uninstall, Status=$Status, Stop=$Stop, Start=$Start"

# --- Status mode (no admin required) ---
if ($Status) {
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($svc) {
        $color = if ($svc.Status -eq 'Running') { 'Green' } else { 'Yellow' }
        Write-Host "Service '$ServiceName': $($svc.Status) (StartType: $($svc.StartType))" -ForegroundColor $color

        if (Test-Path $NssmPath) {
            $nssmStatus = & $NssmPath status $ServiceName 2>&1
            Write-Host "NSSM status: $nssmStatus"
        }
    } else {
        Write-Host "Service '$ServiceName' is not installed." -ForegroundColor Red
    }
    Log "Status check for '$ServiceName': $(if ($svc) { $svc.Status } else { 'Not installed' })" -Level INFO
    exit 0
}

# --- Stop mode ---
if ($Stop) {
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) {
        Write-Host "Service '$ServiceName' is not installed." -ForegroundColor Red
        exit 1
    }
    Log "Stopping service '$ServiceName'..." -Level INFO
    if (Test-Path $NssmPath) {
        & $NssmPath stop $ServiceName 2>&1 | Out-Null
    } else {
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
    }
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    $color = if ($svc.Status -eq 'Stopped') { 'Green' } else { 'Yellow' }
    Write-Host "Service '$ServiceName': $($svc.Status)" -ForegroundColor $color
    Log "Stop command completed. Status: $($svc.Status)" -Level INFO
    exit 0
}

# --- Start mode ---
if ($Start) {
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) {
        Write-Host "Service '$ServiceName' is not installed." -ForegroundColor Red
        exit 1
    }
    Log "Starting service '$ServiceName'..." -Level INFO
    if (Test-Path $NssmPath) {
        & $NssmPath start $ServiceName 2>&1 | Out-Null
    } else {
        Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
    }
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    $color = if ($svc.Status -eq 'Running') { 'Green' } else { 'Yellow' }
    Write-Host "Service '$ServiceName': $($svc.Status)" -ForegroundColor $color
    Log "Start command completed. Status: $($svc.Status)" -Level INFO
    exit 0
}

# --- No action specified: show help ---
if (-not $Install -and -not $Uninstall -and -not $Status -and -not $Stop -and -not $Start) {
    Show-Help
    exit 0
}

# --- Validate: only one action switch at a time ---
$specifiedActions = @('Install', 'Uninstall', 'Status', 'Stop', 'Start') |
    Where-Object { $PSBoundParameters.ContainsKey($_) }
if ($specifiedActions.Count -gt 1) {
    Write-Host "ERROR: Only one action switch may be specified at a time. Got: -$($specifiedActions -join ', -')" -ForegroundColor Red
    Show-Help
    exit 1
}

# --- Validate: install-only options require -Install ---
$installOnlyParams = @('DisplayName', 'ModelPath', 'OllamaPath')
$badInstallParams = $installOnlyParams | Where-Object { $PSBoundParameters.ContainsKey($_) }
if ($badInstallParams.Count -gt 0 -and -not $Install) {
    Write-Host "ERROR: -$($badInstallParams -join ', -') can only be used with -Install." -ForegroundColor Red
    Show-Help
    exit 1
}

# --- Validate: -Port requires -Install or -Uninstall ---
if ($PSBoundParameters.ContainsKey('Port') -and -not $Install -and -not $Uninstall) {
    Write-Host "ERROR: -Port can only be used with -Install or -Uninstall." -ForegroundColor Red
    Show-Help
    exit 1
}

# --- Validate: -ServiceName is the only option valid with -Status, -Stop, -Start ---
$managementOnlyAllowed = @('ServiceName', 'Status', 'Stop', 'Start', 'Silent', 'NssmPath')
if ($Status -or $Stop -or $Start) {
    $extraParams = $PSBoundParameters.Keys | Where-Object { $managementOnlyAllowed -notcontains $_ }
    if ($extraParams) {
        Write-Host "ERROR: -$($extraParams -join ', -') cannot be used with -$($specifiedActions -join '/-')." -ForegroundColor Red
        Show-Help
        exit 1
    }
}

# --- Check admin ---
function Assert-Admin {
    try {
        $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    }
    catch {
        Log "ERROR: Failed to retrieve current Windows identity: $_" -Level ERROR
        throw
    }
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Log "ERROR: Script must be run as Administrator." -Level ERROR
        throw "Run this script in an elevated PowerShell session."
    }
}
Assert-Admin

# --- Uninstall mode ---
if ($Uninstall) {
    Log "Uninstall mode requested." -Level INFO

    $serviceRemoved = $false

    # Try NSSM first if available
    if (Test-Path $NssmPath) {
        try {
            $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
            if ($svc) {
                Log "Stopping and removing service '$ServiceName' using NSSM..." -Level INFO

                # Stop service with retry logic
                $stopAttempts = 0
                $maxStopAttempts = 3
                do {
                    $stopAttempts++
                    & $NssmPath stop $ServiceName 2>&1 | Out-Null
                    $exitCode = $LASTEXITCODE

                    if ($exitCode -eq 0) {
                        break
                    }

                    Log "NSSM stop attempt $stopAttempts failed (exit code: $exitCode), retrying..." -Level WARNING
                    Start-Sleep -Seconds 2
                } while ($stopAttempts -lt $maxStopAttempts)

                # Wait for service to fully stop
                $serviceStopped = $false
                $stopWaitAttempts = 0
                $maxStopWaitAttempts = 10
                do {
                    $stopWaitAttempts++
                    $currentService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                    if (-not $currentService -or $currentService.Status -eq 'Stopped') {
                        $serviceStopped = $true
                        break
                    }
                    Start-Sleep -Seconds 1
                } while ($stopWaitAttempts -lt $maxStopWaitAttempts)

                if (-not $serviceStopped) {
                    Log "Warning: Service may not have fully stopped, proceeding with removal anyway" -Level WARNING
                }

                # Remove service
                & $NssmPath remove $ServiceName confirm 2>&1 | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Log "Service '$ServiceName' removed successfully via NSSM." -Level INFO
                    $serviceRemoved = $true
                } else {
                    Log "NSSM remove failed (exit code: $LASTEXITCODE), will try manual removal" -Level WARNING
                }
            } else {
                Log "Service '$ServiceName' not found." -Level INFO
            }
        }
        catch {
            Log "Error during NSSM service removal: $_" -Level ERROR
        }
    } else {
        Log "NSSM not found, attempting manual service removal..." -Level INFO
    }

    # Fallback to manual removal if NSSM failed or wasn't available
    if (-not $serviceRemoved) {
        try {
            $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
            if ($svc) {
                Log "Attempting manual service removal..." -Level INFO

                # Force stop service
                Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue

                # Wait for service to stop
                $manualStopWait = 0
                $maxManualStopWait = 5
                do {
                    $manualStopWait++
                    $currentService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                    if (-not $currentService -or $currentService.Status -eq 'Stopped') {
                        break
                    }
                    Start-Sleep -Seconds 1
                } while ($manualStopWait -lt $maxManualStopWait)

                # Delete service
                $deleteResult = sc.exe delete $ServiceName 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Log "Service '$ServiceName' removed manually." -Level INFO
                } else {
                    Log "Manual service removal failed: $deleteResult" -Level ERROR
                }
            }
        }
        catch {
            Log "Error during manual service removal: $_" -Level ERROR
        }
    }

    # Remove firewall rule
    $fwRuleName = "Ollama Port $Port"
    try {
        $rule = Get-NetFirewallRule -DisplayName $fwRuleName -ErrorAction SilentlyContinue
        if ($rule) {
            Log "Removing firewall rule '$fwRuleName'..." -Level INFO
            Remove-NetFirewallRule -DisplayName $fwRuleName -ErrorAction Stop
            Log "Firewall rule removed successfully." -Level INFO
        } else {
            Log "Firewall rule '$fwRuleName' not found." -Level INFO
        }
    }
    catch {
        Log "Warning: Failed to remove firewall rule: $_" -Level WARNING
    }

    Log "=== Uninstall completed ==="
    if (-not $Silent) {
        Write-Host "Ollama service uninstalled. Log: $LogFile"
    }
    exit 0
}

# --- Install mode ---
if ($Install) {

# --- Download NSSM if not present ---
if (-not (Test-Path $NssmPath)) {
    Log "NSSM not found. Downloading..." -Level INFO

    # Validate NSSM download URL
    $nssmZipUrl = "https://nssm.cc/release/nssm-2.24.zip"
    $nssmZipPath = $null
    $nssmExtractPath = $null

    try {
        $nssmZipPath = Join-Path -Path $env:TEMP -ChildPath "nssm.zip"
        $nssmExtractPath = Join-Path -Path $env:TEMP -ChildPath "nssm_temp"

        # Clean up any existing temp files
        if (Test-Path $nssmZipPath) {
            Remove-Item -Path $nssmZipPath -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path $nssmExtractPath) {
            Remove-Item -Path $nssmExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        }

        # Download with security and retry logic
        Log "Downloading NSSM from $nssmZipUrl..." -Level INFO
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $downloadAttempts = 0
        $maxDownloadAttempts = 3
        do {
            $downloadAttempts++
            try {
                Invoke-WebRequest -Uri $nssmZipUrl -OutFile $nssmZipPath -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
                break
            }
            catch {
                if ($downloadAttempts -ge $maxDownloadAttempts) {
                    throw
                }
                Log "Download attempt $downloadAttempts failed, retrying..." -Level WARNING
                Start-Sleep -Seconds 2
            }
        } while ($downloadAttempts -lt $maxDownloadAttempts)

        # Verify download
        if (-not (Test-Path $nssmZipPath) -or (Get-Item $nssmZipPath).Length -lt 1000) {
            throw "Downloaded file appears to be invalid or corrupted"
        }

        # Extract
        Log "Extracting NSSM..." -Level INFO
        Expand-Archive -Path $nssmZipPath -DestinationPath $nssmExtractPath -Force -ErrorAction Stop

        # Determine architecture and find appropriate nssm.exe with improved detection
        $arch = if ([Environment]::Is64BitOperatingSystem) { "win64" } else { "win32" }
        Log "Detected architecture: $arch" -Level DEBUG

        $nssmExeSource = Get-ChildItem -Path $nssmExtractPath -Recurse -Filter "nssm.exe" |
            Where-Object { $_.FullName -like "*\$arch\*" -and $_.Exists } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if ($nssmExeSource) {
            # Verify the executable
            try {
                $fileInfo = Get-Item $nssmExeSource.FullName
                if ($fileInfo.Length -gt 0) {
                    $nssmInstallDir = Split-Path $NssmPath -Parent
                    if (-not (Test-Path $nssmInstallDir)) {
                        New-Item -Path $nssmInstallDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                        Log "Created NSSM install directory: $nssmInstallDir" -Level DEBUG
                    }
                    Copy-Item -Path $nssmExeSource.FullName -Destination $NssmPath -Force -ErrorAction Stop

                    # Set secure permissions - use typed enums; for a file (not dir)
                    # InheritanceFlags and PropagationFlags must be None
                    $acl = Get-Acl $NssmPath
                    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "Administrators",
                        [System.Security.AccessControl.FileSystemRights]::FullControl,
                        [System.Security.AccessControl.InheritanceFlags]::None,
                        [System.Security.AccessControl.PropagationFlags]::None,
                        [System.Security.AccessControl.AccessControlType]::Allow)
                    $acl.SetAccessRule($accessRule)
                    Set-Acl $NssmPath $acl

                    Log "NSSM copied and secured at $NssmPath" -Level INFO
                } else {
                    throw "Found nssm.exe but file appears to be empty"
                }
            }
            catch {
                throw "Failed to copy or secure NSSM executable: $_"
            }
        } else {
            throw "Could not find valid nssm.exe for $arch architecture in extracted files"
        }

        Log "NSSM downloaded and extracted successfully." -Level INFO
    }
    catch {
        Log "ERROR downloading NSSM: $_" -Level ERROR
        throw "Failed to download NSSM. Please download manually from https://nssm.cc/download and place nssm.exe at '$NssmPath' (or use -NssmPath to specify a custom location)."
    }
    finally {
        # Cleanup temporary files
        if ($nssmZipPath -and (Test-Path $nssmZipPath)) {
            Remove-Item -Path $nssmZipPath -Force -ErrorAction SilentlyContinue
        }
        if ($nssmExtractPath -and (Test-Path $nssmExtractPath)) {
            Remove-Item -Path $nssmExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

if (-not (Test-Path $NssmPath)) {
    Log "ERROR: NSSM not found at $NssmPath" -Level ERROR
    throw "NSSM is required but not found."
}

Log "Using NSSM at: $NssmPath"

# --- Detect Ollama path ---
if (-not $OllamaPath) {
    $PossiblePaths = @(
        "$env:LOCALAPPDATA\Programs\Ollama\ollama.exe",
        "$env:PROGRAMFILES\Ollama\ollama.exe",
        "${env:PROGRAMFILES(X86)}\Ollama\ollama.exe",
        "$env:USERPROFILE\ollama.exe",
        ".\ollama.exe",
        "ollama.exe"  # Check PATH
    )

    Log "Searching for Ollama executable..." -Level DEBUG

    foreach ($path in $PossiblePaths) {
        if (Test-Path $path) {
            $OllamaPath = $path
            Log "Found Ollama at: $OllamaPath" -Level DEBUG
            break
        }
    }

    # If still not found, try to find it in PATH using where.exe
    if (-not $OllamaPath) {
        try {
            $pathResult = where.exe ollama.exe 2>$null
            if ($pathResult) {
                $OllamaPath = $pathResult[0].Trim()
                Log "Found Ollama in PATH: $OllamaPath" -Level DEBUG
            }
        }
        catch {
            Log "Failed to search PATH for ollama.exe: $_" -Level DEBUG
        }
    }
}

if (-not $OllamaPath -or -not (Test-Path $OllamaPath)) {
    Log "ERROR: Could not find ollama.exe. Checked standard locations and PATH." -Level ERROR
    throw "Ollama executable not found. Install Ollama or specify -OllamaPath parameter."
}

# Verify Ollama executable
try {
    $ollamaInfo = Get-Item $OllamaPath
    if ($ollamaInfo.Length -lt 1000) {
        Log "Warning: Ollama executable appears to be unusually small ($($ollamaInfo.Length) bytes)" -Level WARNING
    }
}
catch {
    Log "Warning: Could not verify Ollama executable properties: $_" -Level WARNING
}

Log "Using Ollama executable: $OllamaPath" -Level INFO

# --- Remove existing service if present ---
$existing = $null
try {
    $existing = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
}
catch {
    Log "Error checking for existing service: $_" -Level DEBUG
}

if ($existing) {
    Log "Existing service '$ServiceName' found. Removing..." -Level INFO

    try {
        # Stop service with retry logic
        $stopAttempts = 0
        $maxStopAttempts = 3
        do {
            $stopAttempts++
            & $NssmPath stop $ServiceName 2>&1 | Out-Null
            $exitCode = $LASTEXITCODE

            if ($exitCode -eq 0) {
                break
            }

            Log "NSSM stop attempt $stopAttempts failed (exit code: $exitCode), retrying..." -Level WARNING
            Start-Sleep -Seconds 2
        } while ($stopAttempts -lt $maxStopAttempts)

        # Wait for service to fully stop
        $serviceStopped = $false
        $stopWaitAttempts = 0
        $maxStopWaitAttempts = 10
        do {
            $stopWaitAttempts++
            $currentService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
            if (-not $currentService -or $currentService.Status -eq 'Stopped') {
                $serviceStopped = $true
                break
            }
            Start-Sleep -Seconds 1
        } while ($stopWaitAttempts -lt $maxStopWaitAttempts)

        if (-not $serviceStopped) {
            Log "Warning: Service may not have fully stopped, proceeding with removal anyway" -Level WARNING
        }

        # Remove service
        & $NssmPath remove $ServiceName confirm 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Log "Old service removed successfully." -Level INFO
        } else {
            Log "Warning: NSSM remove failed (exit code: $LASTEXITCODE)" -Level WARNING
        }

        Start-Sleep -Seconds 2
    }
    catch {
        Log "Error removing existing service: $_" -Level ERROR
        throw "Failed to remove existing service. Please remove it manually and retry."
    }
}

# --- Build environment variables ---
$envVars = @{
    "OLLAMA_HOST" = "0.0.0.0:$Port"
}

if ($ModelPath) {
    $envVars["OLLAMA_MODELS"] = $ModelPath
    Log "Using custom model path: $ModelPath"
}

# --- Install service with NSSM ---
try {
    Log "Installing service '$ServiceName' with NSSM..." -Level INFO

    # Install service
    $installResult = & $NssmPath install $ServiceName $OllamaPath serve 2>&1
    $installExitCode = $LASTEXITCODE

    if ($installExitCode -ne 0) {
        Log "NSSM install failed with exit code $installExitCode" -Level ERROR
        Log "NSSM install output: $installResult" -Level ERROR
        throw "NSSM install failed with exit code $installExitCode"
    }

    Log "NSSM install completed successfully." -Level DEBUG

    # Configure service
    $configSteps = @(
        { & $NssmPath set $ServiceName DisplayName $DisplayName 2>&1 | Out-Null },
        { & $NssmPath set $ServiceName Description "Ollama AI model server running on port $Port" 2>&1 | Out-Null },
        { & $NssmPath set $ServiceName Start SERVICE_AUTO_START 2>&1 | Out-Null }
    )

    foreach ($step in $configSteps) {
        try {
            & $step
            if ($LASTEXITCODE -ne 0) {
                throw "Configuration step failed with exit code $LASTEXITCODE"
            }
        }
        catch {
            Log "Service configuration failed: $_" -Level ERROR
            throw "Failed to configure service: $_"
        }
    }

    # Set environment variables - all in a single nssm call to avoid each call
    # overwriting the previous AppEnvironmentExtra value
    $validEnvPairs = @()
    foreach ($key in $envVars.Keys) {
        $value = $envVars[$key]
        if ([string]::IsNullOrWhiteSpace($key) -or [string]::IsNullOrWhiteSpace($value)) {
            Log "Warning: Skipping invalid environment variable: key='$key', value='$value'" -Level WARNING
            continue
        }
        $validEnvPairs += "$key=$value"
        Log "Queued environment variable: $key=$value" -Level DEBUG
    }

    if ($validEnvPairs.Count -gt 0) {
        try {
            $envResult = & $NssmPath set $ServiceName AppEnvironmentExtra @validEnvPairs 2>&1
            if ($LASTEXITCODE -eq 0) {
                Log "Set $($validEnvPairs.Count) environment variable(s) successfully." -Level DEBUG
            } else {
                Log "Failed to set environment variables: $envResult" -Level ERROR
                throw "Failed to set AppEnvironmentExtra (exit code $LASTEXITCODE)"
            }
        }
        catch {
            Log "Error setting environment variables: $_" -Level ERROR
            throw "Failed to set environment variables: $_"
        }
    }

    # Configure stdout/stderr logging
    $logDir = "$DataDir\logs"
    if (-not (Test-Path $logDir)) {
        try {
            New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Log "Created log directory: $logDir" -Level DEBUG
        }
        catch {
            Log "Failed to create log directory: $_" -Level ERROR
            throw "Failed to create log directory: $_"
        }
    }

    $stdoutLog = Join-Path -Path $logDir -ChildPath "ollama-stdout.log"
    $stderrLog = Join-Path -Path $logDir -ChildPath "ollama-stderr.log"

    $logConfigSteps = @(
        { & $NssmPath set $ServiceName AppStdout $stdoutLog 2>&1 | Out-Null },
        { & $NssmPath set $ServiceName AppStderr $stderrLog 2>&1 | Out-Null },
        { & $NssmPath set $ServiceName AppRotateFiles 1 2>&1 | Out-Null },
        { & $NssmPath set $ServiceName AppRotateBytes 1048576 2>&1 | Out-Null }
    )

    foreach ($step in $logConfigSteps) {
        try {
            & $step
            if ($LASTEXITCODE -ne 0) {
                throw "Log configuration step failed with exit code $LASTEXITCODE"
            }
        }
        catch {
            Log "Log configuration failed: $_" -Level WARNING
            # Don't fail the entire installation for log configuration issues
        }
    }

    Log "Service '$ServiceName' installed successfully." -Level INFO
}
catch {
    Log "ERROR installing service: $_" -Level ERROR
    throw "Failed to install service with NSSM: $_"
}

# --- Configure firewall rule ---
try {
    $fwRuleName = "Ollama Port $Port"
    $existingRule = Get-NetFirewallRule -DisplayName $fwRuleName -ErrorAction SilentlyContinue

    if ($existingRule) {
        Log "Firewall rule '$fwRuleName' already exists." -Level INFO
    } else {
        Log "Creating firewall rule '$fwRuleName' for TCP port $Port..." -Level INFO

        $firewallParams = @{
            DisplayName = $fwRuleName
            Direction = "Inbound"
            Action = "Allow"
            Protocol = "TCP"
            LocalPort = $Port
            Profile = "Any"
            Description = "Allow inbound connections to Ollama service on port $Port"
            ErrorAction = "Stop"
        }

        New-NetFirewallRule @firewallParams | Out-Null
        Log "Firewall rule created successfully." -Level INFO
    }
}
catch {
    Log "Warning: Failed to create firewall rule: $_" -Level WARNING
    Log "The service will still work, but may require manual firewall configuration." -Level INFO
}

# --- Start service ---
try {
    Log "Starting service '$ServiceName'..." -Level INFO

    $startResult = & $NssmPath start $ServiceName 2>&1
    $startExitCode = $LASTEXITCODE

    if ($startExitCode -eq 0) {
        Log "NSSM start command executed successfully." -Level DEBUG

        # Wait for service to start with timeout and retry logic
        $serviceStarted = $false
        $startWaitAttempts = 0
        $maxStartWaitAttempts = 30  # 30 seconds total timeout

        do {
            $startWaitAttempts++
            Start-Sleep -Seconds 1

            try {
                $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                if ($svc -and $svc.Status -eq 'Running') {
                    $serviceStarted = $true
                    Log "Service '$ServiceName' started successfully after $startWaitAttempts seconds." -Level INFO
                    break
                } elseif ($svc) {
                    Log "Service status after $startWaitAttempts seconds: $($svc.Status)" -Level DEBUG
                } else {
                    Log "Service not found after $startWaitAttempts seconds" -Level DEBUG
                }
            }
            catch {
                Log "Error checking service status at attempt ${startWaitAttempts}: $_" -Level DEBUG
            }

            if ($startWaitAttempts -ge $maxStartWaitAttempts) {
                Log "Service startup timeout after $maxStartWaitAttempts seconds" -Level WARNING
                break
            }
        } while ($startWaitAttempts -lt $maxStartWaitAttempts)

        if (-not $serviceStarted) {
            # Get final service status for reporting
            try {
                $finalSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                if ($finalSvc) {
                    Log "Warning: Service failed to start. Final status: $($finalSvc.Status)" -Level WARNING
                    Log "NSSM start output: $startResult" -Level DEBUG
                } else {
                    Log "Warning: Service not found after startup attempt" -Level WARNING
                }
            }
            catch {
                Log "Warning: Could not determine final service status: $_" -Level WARNING
            }

            # Don't throw error for startup failure, just warn
            Log "Service was installed but failed to start. Check logs in $logDir and the Windows Event Viewer." -Level WARNING
        }
    } else {
        Log "NSSM start failed with exit code $startExitCode" -Level ERROR
        Log "NSSM start output: $startResult" -Level ERROR
        throw "NSSM start failed with exit code $startExitCode"
    }
}
catch {
    Log "ERROR starting service: $_" -Level ERROR
    throw "Service installed but failed to start: $_. Check logs in $logDir and the Windows Event Viewer."
}

Log "=== Ollama Service Installer completed successfully ==="

if (-not $Silent) {
    $scriptName = $MyInvocation.MyCommand.Name
    Write-Host ""
    Write-Host "Ollama service installed and running on port $Port." -ForegroundColor Green
    Write-Host "Service logs: $logDir"
    Write-Host "Installation log: $LogFile"
    Write-Host ""
    Write-Host "To manage the service:"
    Write-Host "  Start:     .\$scriptName -Start"
    Write-Host "  Stop:      .\$scriptName -Stop"
    Write-Host "  Status:    .\$scriptName -Status"
    Write-Host "  Uninstall: .\$scriptName -Uninstall"
}

} # end if ($Install)
