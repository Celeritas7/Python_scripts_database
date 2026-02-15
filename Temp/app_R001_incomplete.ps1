# ========= Config =========
$ToolRoot = "C:\tools\media-tools"   # your toolbox root
$Python   = "python"                 # or full path to python.exe if needed

function Select-FromList {
    param(
        [string]$Title,
        [string[]]$Items
    )
    if (-not $Items -or $Items.Count -eq 0) {
        Write-Host "No items found for: $Title" -ForegroundColor Yellow
        return $null
    }

    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan
    for ($i=0; $i -lt $Items.Count; $i++) {
        "{0,2}) {1}" -f ($i+1), $Items[$i] | Write-Host
    }
    Write-Host ""

    while ($true) {
        $raw = Read-Host "Select 1-$($Items.Count) (or q to quit)"
        if ($raw -eq "q") { return $null }
        if ([int]::TryParse($raw, [ref]$n) -and $n -ge 1 -and $n -le $Items.Count) {
            return $Items[$n-1]
        }
        Write-Host "Invalid input." -ForegroundColor Red
    }
}

# ========= Step 1: pick category =========
$categories = Get-ChildItem -Path $ToolRoot -Directory |
              Where-Object { $_.Name -notmatch '^(Old|\.git)$' } |
              Select-Object -ExpandProperty Name |
              Sort-Object

$category = Select-FromList -Title "Choose a category (Toolbox: $ToolRoot)" -Items $categories
if (-not $category) { return }

$catPath = Join-Path $ToolRoot $category

# ========= Step 2: pick task (script) =========
# Accept python or powershell tools as tasks
$tasks = Get-ChildItem -Path $catPath -File |
         Where-Object { $_.Extension -in ".py", ".ps1" } |
         Select-Object -ExpandProperty Name |
         Sort-Object

$task = Select-FromList -Title "Choose a task in [$category]" -Items $tasks
if (-not $task) { return }

$taskPath = Join-Path $catPath $task

# ========= Optional: ask for extra args =========
Write-Host ""
Write-Host "Current working folder: $pwd" -ForegroundColor Green
$extraArgs = Read-Host "Optional arguments to pass to the tool (press Enter for none)"

# ========= Run =========
Write-Host ""
Write-Host "Running: $taskPath" -ForegroundColor Cyan

if ($taskPath.EndsWith(".py")) {
    if ([string]::IsNullOrWhiteSpace($extraArgs)) {
        & $Python $taskPath
    } else {
        & $Python $taskPath @($extraArgs -split '\s+')
    }
}
elseif ($taskPath.EndsWith(".ps1")) {
    if ([string]::IsNullOrWhiteSpace($extraArgs)) {
        & powershell -ExecutionPolicy Bypass -File $taskPath
    } else {
        & powershell -ExecutionPolicy Bypass -File $taskPath @($extraArgs -split '\s+')
    }
}
