# MLS Procedures Auto-Sync Script
$ErrorActionPreference = "SilentlyContinue"

$source = "\\pfbafileserver\File Server\Admin Ops\Research\MLS\Procedures"
$dest = "C:\Users\karthik.padmanathan\mls_docs_portal\sample_docs"
$repo = "C:\Users\karthik.padmanathan\mls_docs_portal"
$gitExe = "$HOME\AppData\Local\Programs\Git\cmd\git.exe"
$logFile = "$repo\sync.log"

function Write-Log($msg) {
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $logFile -Value "[$timestamp] $msg"
}

# 1. Check if network file server is accessible
if (-not (Test-Path -LiteralPath $source)) {
    # Exit silently if not on corporate network/VPN
    exit 0
}

# 2. Check if git exists
if (-not (Test-Path -LiteralPath $gitExe)) {
    Write-Log "Git executable not found at $gitExe"
    exit 1
}

# 3. Robocopy files from network share to local sample_docs
# Robocopy exit codes: 0 = no changes, 1 = files copied successfully
$robocopyArgs = @(
    $source,
    $dest,
    "/MIR",
    "/FFT",
    "/Z",
    "/NP",
    "/NDL",
    "/NJH",
    "/NJS",
    "/XF", "~$*", "~*", "*.tmp"
)

& robocopy.exe @robocopyArgs
$roboExit = $LASTEXITCODE

# 4. Check git status to see if anything changed
Set-Location $repo
$status = & $gitExe status --porcelain "sample_docs"

if ($status) {
    Write-Log "New/updated documents detected. Committing and pushing to GitHub..."
    & $gitExe add sample_docs/
    & $gitExe commit -m "Auto-sync: updated procedures from network file server"
    $pushOutput = & $gitExe push origin main 2>&1
    Write-Log "Push complete: $pushOutput"
} else {
    Write-Log "Sync check OK: No new changes on file server."
    if (Test-Path $logFile) {
        $lines = Get-Content $logFile -Tail 20
        Set-Content -Path $logFile -Value $lines
    }
}
