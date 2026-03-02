# D4-Ollama-Win-Service

Windows Service installation and management for Ollama using NSSM (Non-Sucking Service Manager).

## Overview

`ollama.exe` does not implement the Windows Service Control Manager (SCM) interface and cannot run as a native Windows service. NSSM acts as the service binary, spawning and managing `ollama.exe` and handling all SCM communication.

**Native service error without NSSM:** `%%1053 - The service did not respond to the start or control request in a timely fashion`

## Files

- `scripts/win_ollama_service_install_nssm.ps1` — PowerShell service manager script

## Quick Start

```powershell
# Run as Administrator in PowerShell terminal
.\scripts\win_ollama_service_install_nssm.ps1 -Install
```

Running without arguments displays built-in help.

## Usage

| Command | Requires Admin | Description |
|---|---|---|
| `-Install` | Yes | Download NSSM (if needed), install and start the service |
| `-Uninstall` | Yes | Stop and remove the service and firewall rule |
| `-Start` | Yes | Start the installed service |
| `-Stop` | Yes | Stop the running service |
| `-Status` | No | Display current service state |

### Install

```powershell
# Default install (service name: ollama, port: 11434)
.\scripts\win_ollama_service_install_nssm.ps1 -Install

# Custom port
.\scripts\win_ollama_service_install_nssm.ps1 -Install -Port 8080

# Custom model storage directory
.\scripts\win_ollama_service_install_nssm.ps1 -Install -ModelPath "D:\OllamaModels"

# Custom ollama.exe path (auto-detected if omitted)
.\scripts\win_ollama_service_install_nssm.ps1 -Install -OllamaPath "C:\Custom\ollama.exe"

# All options
.\scripts\win_ollama_service_install_nssm.ps1 -Install -ServiceName "myollama" -Port 8080 -ModelPath "D:\OllamaModels"
```

### Start / Stop / Status

```powershell
# Start the service (requires Admin)
.\scripts\win_ollama_service_install_nssm.ps1 -Start

# Stop the service (requires Admin)
.\scripts\win_ollama_service_install_nssm.ps1 -Stop

# Check service status (no Admin required)
.\scripts\win_ollama_service_install_nssm.ps1 -Status
```

### Uninstall

```powershell
# Run as Administrator in PowerShell terminal
.\scripts\win_ollama_service_install_nssm.ps1 -Uninstall

# Uninstall a custom-named service
.\scripts\win_ollama_service_install_nssm.ps1 -Uninstall -ServiceName "myollama"
```

> **Note:** Uninstall removes the Windows service and firewall rule. `C:\Program Files\NSSM\nssm.exe` is intentionally left in place — NSSM may be used by other services. Remove it manually if no longer needed.

## Parameters

| Parameter | Used with | Default | Description |
|---|---|---|---|
| `-Install` | — | — | Install and start the service |
| `-Uninstall` | — | — | Stop and remove the service |
| `-Start` | — | — | Start the installed service |
| `-Stop` | — | — | Stop the running service |
| `-Status` | — | — | Display current service status |
| `-ServiceName` | all | `ollama` | Windows service name |
| `-DisplayName` | `-Install` | `Ollama Service` | Display name shown in Services |
| `-Port` | `-Install`, `-Uninstall` | `11434` | Listening port (sets `OLLAMA_HOST`) |
| `-ModelPath` | `-Install` | *(auto)* | Custom model storage directory |
| `-OllamaPath` | `-Install` | *(auto-detected)* | Full path to `ollama.exe` |
| `-NssmPath` | all | `C:\Program Files\NSSM\nssm.exe` | Path to `nssm.exe` |
| `-Silent` | all | — | Suppress console output |

## Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1+ (run as Administrator for install/uninstall/start/stop)
- Ollama installed
- Internet connection (for automatic NSSM download, only if `C:\Program Files\NSSM\nssm.exe` is not already present)

## Features

- Automatic NSSM download → installed to `C:\Program Files\NSSM\nssm.exe` (stable across repo moves)
- Auto-detects `ollama.exe` installation path
- Configurable port, service name, model path, custom NSSM path (`-NssmPath`)
- Inbound firewall rule created automatically
- Logs written to `C:\ProgramData\Ollama\` with 1 MB rotation
- Retry logic for service start/stop operations
- Built-in help — running without arguments shows usage
- Argument validation — invalid combinations show help and exit

## NSSM Setup

### Option 1: Automatic Download (Recommended)

If `nssm.exe` is not found at `C:\Program Files\NSSM\nssm.exe`, `-Install` automatically downloads it from `https://nssm.cc/release/nssm-2.24.zip` and installs it there.

### Option 2: Manual

1. Download NSSM from https://nssm.cc/download
2. Place `nssm.exe` at `C:\Program Files\NSSM\nssm.exe`
3. Run `.\scripts\win_ollama_service_install_nssm.ps1 -Install`

### Custom NSSM location

```powershell
.\scripts\win_ollama_service_install_nssm.ps1 -Install -NssmPath "D:\Tools\nssm.exe"
```

## File Locations

| File | Path |
|---|---|
| `nssm.exe` | `C:\Program Files\NSSM\nssm.exe` |
| Install log | `C:\ProgramData\Ollama\OllamaServiceInstall.log` |
| Service stdout | `C:\ProgramData\Ollama\logs\ollama-stdout.log` |
| Service stderr | `C:\ProgramData\Ollama\logs\ollama-stderr.log` |

Service logs rotate automatically at 1 MB.

```powershell
# View recent service output
Get-Content "$env:ProgramData\Ollama\logs\ollama-stdout.log" -Tail 50
Get-Content "$env:ProgramData\Ollama\logs\ollama-stderr.log" -Tail 50

# View install log
Get-Content "$env:ProgramData\Ollama\OllamaServiceInstall.log" -Tail 50
```

## How It Works

1. **NSSM wraps ollama.exe** — NSSM runs as the Windows service and manages the `ollama serve` process
2. **Environment variables** — Sets `OLLAMA_HOST=0.0.0.0:<port>` and optionally `OLLAMA_MODELS`
3. **Auto-restart** — NSSM restarts `ollama.exe` automatically if it crashes
4. **Log rotation** — Logs rotate at 1 MB automatically
5. **Firewall rule** — An inbound TCP rule is created for the configured port

- Ollama service will always listen to **0.0.0.0**, that is to all network interfaces, and the port 11434 or the port specified in the install command.

- Models are stored by default in **`C:\Windows\system32\config\systemprofile\.ollama\models`** or the path specified in the install command.

## Troubleshooting

### Service won't start

```powershell
.\scripts\win_ollama_service_install_nssm.ps1 -Status
Get-Content "$env:ProgramData\Ollama\logs\ollama-stderr.log" -Tail 50
```

### Port already in use

```powershell
# Find what is using the port
netstat -ano | findstr :11434

# Stop any stray Ollama processes
Get-Process ollama -ErrorAction SilentlyContinue | Stop-Process -Force

# Restart the service
.\scripts\win_ollama_service_install_nssm.ps1 -Stop
.\scripts\win_ollama_service_install_nssm.ps1 -Start
```

### Move or re-clone the repo without breaking the service

Because `nssm.exe` is installed at **`C:\Program Files\NSSM\nssm.exe`** (not next to the script), the running service is unaffected if you move, delete, or re-clone the repository. The service keeps running — only the management script location changes.

```powershell
# After re-cloning to a new path, manage the service from the new location
.\new\path\d4-ollama-win-service\scripts\win_ollama_service_install_nssm.ps1 -Status
```

### Verify Ollama is responding

```powershell
Invoke-WebRequest -Uri "http://localhost:11434/api/tags" -Method GET
```

## Comparison: Native Service vs NSSM

| Feature | Native Windows Service | NSSM Wrapper |
|---|---|---|
| Works with ollama.exe | ❌ Times out | ✅ Yes |
| Auto-restart on crash | ⚠️ Limited | ✅ Full support |
| Log capture + rotation | ❌ Manual | ✅ Automatic |
| Environment variables | ⚠️ Registry edits | ✅ Built-in |
| Easy configuration | ⚠️ Complex | ✅ Script parameters |

## Additional Resources

- NSSM Documentation: https://nssm.cc/usage
- Ollama Documentation: https://github.com/ollama/ollama

## License

MIT License - see LICENSE file for details.
