param(
    [string]$RemoteName = "auto",
    [string]$MountPoint = "H:",
    [string]$ProfileDir = "Default",
    [switch]$ForceRefresh,
    [switch]$DryRun,
    [switch]$NonInteractive,
    [switch]$Headless,
    [string]$LogFile = "",
    [int]$MountCheckTimeoutSeconds = 8
)

$ErrorActionPreference = "Stop"
$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonScript = Join-Path $BaseDir "refresh_sharepoint_rclone_cookies.py"
$VenvPython = Join-Path $BaseDir "venv\Scripts\python.exe"
$Script:ResolvedLogFile = if ($LogFile) { $LogFile } else { Join-Path $BaseDir "logs\rclone-set.log" }
$Script:LockFile = Join-Path $BaseDir ".rclone-set.lock"
$Script:LockStream = $null

# Writes a timestamped message to the console and the launcher log file.
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Write-Host $line
    $logDir = Split-Path -Parent $Script:ResolvedLogFile
    if ($logDir) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    Add-Content -LiteralPath $Script:ResolvedLogFile -Value $line
}

# Prevents multiple launcher instances from running at the same time.
function Enter-ScriptLock {
    try {
        $lockDir = Split-Path -Parent $Script:LockFile
        if ($lockDir) {
            New-Item -ItemType Directory -Path $lockDir -Force | Out-Null
        }
        $Script:LockStream = [System.IO.File]::Open($Script:LockFile, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $Script:LockStream.SetLength(0)
        $writer = New-Object System.IO.StreamWriter($Script:LockStream)
        $writer.AutoFlush = $true
        $writer.WriteLine("pid=$PID")
        $writer.WriteLine("time=$(Get-Date -Format o)")
        $writer.Flush()
        $Script:LockStream.Position = 0
    }
    catch {
        throw "Another launcher instance is already running. Lock file: $Script:LockFile"
    }
}

# Releases the launcher lock file handle.
function Exit-ScriptLock {
    if ($Script:LockStream) {
        $Script:LockStream.Dispose()
        $Script:LockStream = $null
    }
}

# Locates rclone.exe from PATH or a Scoop installation.
function Resolve-RcloneExe {
    $cmd = Get-Command rclone -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    if ($env:SCOOP) {
        $candidate = Join-Path $env:SCOOP "apps\rclone\current\rclone.exe"
        if (Test-Path $candidate) { return $candidate }
    }

    $defaultScoop = Join-Path $HOME "scoop\apps\rclone\current\rclone.exe"
    if (Test-Path $defaultScoop) { return $defaultScoop }

    throw "Could not find rclone.exe"
}

# Reads the rclone config location by parsing the output of 'rclone config file'.
function Get-RcloneConfigPath {
    param(
        [string]$RcloneExe
    )

    $output = & $RcloneExe config file 2>$null
    if (-not $output) {
        throw "Could not determine rclone.conf path from rclone."
    }

    $lines = @($output | Where-Object { $_ -and $_.Trim() })
    [array]::Reverse($lines)
    foreach ($line in $lines) {
        $candidate = $line.Trim().Trim('"')
        if ($candidate -match 'stored at:\s*(.+)$') {
            $candidate = $Matches[1].Trim().Trim('"')
        }
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    throw "Could not resolve rclone.conf path from rclone output."
}

# Resolves the target remote name.
# If 'auto' is used, it selects the only SharePoint WebDAV remote found in rclone.conf.
function Resolve-RemoteName {
    param(
        [string]$RemoteName,
        [string]$RcloneExe
    )

    if ($RemoteName -ne "auto") {
        return $RemoteName
    }

    $configPath = Get-RcloneConfigPath -RcloneExe $RcloneExe
    $content = Get-Content -LiteralPath $configPath
    $current = $null
    $type = $null
    $url = $null
    $remoteMatches = New-Object System.Collections.Generic.List[string]

    foreach ($line in $content) {
        if ($line -match '^\[(.+?)\]\s*$') {
            if ($current -and $type -eq 'webdav' -and $url -match 'sharepoint\.com') {
                [void]$remoteMatches.Add($current)
            }
            $current = $Matches[1]
            $type = $null
            $url = $null
            continue
        }

        if (-not $current) {
            continue
        }

        if ($line -match '^\s*type\s*=\s*(.+?)\s*$') {
            $type = $Matches[1].Trim().ToLowerInvariant()
            continue
        }

        if ($line -match '^\s*url\s*=\s*(.+?)\s*$') {
            $url = $Matches[1].Trim()
            continue
        }
    }

    if ($current -and $type -eq 'webdav' -and $url -match 'sharepoint\.com') {
        [void]$remoteMatches.Add($current)
    }

    if ($remoteMatches.Count -eq 1) {
        return $remoteMatches[0]
    }

    if ($remoteMatches.Count -gt 1) {
        throw "Multiple SharePoint WebDAV remotes found in rclone.conf: $($remoteMatches -join ', '). Please specify -RemoteName explicitly."
    }

    throw "Could not auto detect a SharePoint WebDAV remote. Please specify -RemoteName explicitly."
}

# Locates python.exe, preferring the local virtual environment if present.
function Resolve-PythonExe {
    if (Test-Path $VenvPython) {
        return $VenvPython
    }

    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        return $pythonCmd.Source
    }

    throw "Could not find python.exe. Expected venv interpreter at '$VenvPython' or python on PATH."
}

# Checks whether the mount point exists and responds within a short timeout.
# A background job is used so the script does not hang on stale mounts.
function Test-MountPointAccessible {
    param(
        [string]$MountPoint,
        [int]$TimeoutSeconds = 8
    )

    try {
        if (-not (Test-Path $MountPoint)) {
            return $false
        }

        $job = Start-Job -ScriptBlock {
            param($Path)
            Get-ChildItem -LiteralPath $Path -ErrorAction Stop | Select-Object -First 1 | Out-Null
            $true
        } -ArgumentList $MountPoint

        try {
            $completed = Wait-Job -Job $job -Timeout $TimeoutSeconds
            if (-not $completed) {
                Stop-Job -Job $job -Force | Out-Null
                return $false
            }
            $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
            return [bool]$result
        }
        finally {
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null
        }
    }
    catch {
        return $false
    }
}

# Finds rclone mount processes that match the requested remote and mount point.
function Get-RcloneMountProcesses {
    param(
        [string]$RemoteName,
        [string]$MountPoint
    )

    $matched = New-Object System.Collections.Generic.List[object]
    $procs = @(Get-CimInstance Win32_Process -Filter "Name = 'rclone.exe'" -ErrorAction SilentlyContinue)
    foreach ($proc in $procs) {
        $cmd = $proc.CommandLine
        if (-not $cmd) {
            continue
        }

        $hasMountVerb = $cmd -match '(^|\s)mount(\s|$)'
        $hasMountPoint = $cmd -match [regex]::Escape($MountPoint)
        $hasRemote = $true
        if ($RemoteName -ne "auto") {
            $hasRemote = $cmd -match [regex]::Escape("$RemoteName`:")
        }

        if ($hasMountVerb -and $hasMountPoint -and $hasRemote) {
            [void]$matched.Add($proc)
        }
    }

    return @($matched.ToArray())
}

# Stops one or more stale rclone mount processes.
function Stop-RcloneMountProcesses {
    param(
        [object[]]$Processes
    )

    foreach ($proc in $Processes) {
        try {
            Stop-Process -Id $proc.ProcessId -Force -ErrorAction Stop
            Write-Log "Stopped stale rclone process PID $($proc.ProcessId)."
        }
        catch {
            Write-Log "Failed to stop stale rclone process PID $($proc.ProcessId): $($_.Exception.Message)" "WARN"
        }
    }
}

if (-not (Test-Path $PythonScript)) {
    Write-Error "Python script not found: $PythonScript"
    exit 1
}

try {
    Enter-ScriptLock

    $pythonExe = Resolve-PythonExe
    $rcloneExe = Resolve-RcloneExe
    $resolvedRemoteName = Resolve-RemoteName -RemoteName $RemoteName -RcloneExe $rcloneExe

    Write-Log "Using python: $pythonExe"
    Write-Log "Using rclone: $rcloneExe"
    Write-Log "Using remote: $resolvedRemoteName"
    Write-Log "Log file: $Script:ResolvedLogFile"
    Write-Log "Refreshing SharePoint cookies if needed..."

    $pythonArgs = @(
        $PythonScript,
        "--remote", $resolvedRemoteName,
        "--profile-directory", $ProfileDir,
        "--rclone-exe", $rcloneExe,
        "--log-file", $Script:ResolvedLogFile,
        "--lock-file", (Join-Path $BaseDir ".cookie-refresh.lock")
    )

    if ($ForceRefresh) {
        $pythonArgs += "--force-refresh"
    }
    if ($DryRun) {
        $pythonArgs += "--dry-run"
    }
    if ($NonInteractive) {
        $pythonArgs += "--non-interactive"
    }
    if ($Headless) {
        $pythonArgs += "--headless"
    }

    & $pythonExe @pythonArgs
    $pythonExitCode = $LASTEXITCODE
    if ($pythonExitCode -ne 0) {
        throw "Python cookie refresh script failed with exit code $pythonExitCode."
    }

    if ($DryRun) {
        Write-Log "Dry run completed. Skipping mount start."
        exit 0
    }

    $mountProcesses = @(Get-RcloneMountProcesses -RemoteName $resolvedRemoteName -MountPoint $MountPoint)
    $mountAccessible = Test-MountPointAccessible -MountPoint $MountPoint -TimeoutSeconds $MountCheckTimeoutSeconds

    if ($mountProcesses.Count -gt 0 -and $mountAccessible) {
        Write-Log "rclone mount is already running and accessible on $MountPoint"
        exit 0
    }

    if ($mountProcesses.Count -gt 0 -and -not $mountAccessible) {
        Write-Log "Detected stale rclone mount on $MountPoint. Stopping existing process and restarting." "WARN"
        Stop-RcloneMountProcesses -Processes $mountProcesses
        Start-Sleep -Seconds 2
    }

    Write-Log "Starting rclone mount in background..."
    $rcloneArgs = @(
        "mount",
        "${resolvedRemoteName}:",
        $MountPoint,
        "--vfs-cache-mode", "writes"
    )
    Start-Process -FilePath $rcloneExe -ArgumentList $rcloneArgs -WindowStyle Hidden | Out-Null
    Start-Sleep -Seconds 2

    $mountHealthyAfterStart = Test-MountPointAccessible -MountPoint $MountPoint -TimeoutSeconds $MountCheckTimeoutSeconds
    if ($mountHealthyAfterStart) {
        Write-Log "Mount started successfully on $MountPoint"
        exit 0
    }

    throw "rclone mount process was started, but the mount on $MountPoint is not accessible yet. Review the log file for details."
}
finally {
    Exit-ScriptLock
}