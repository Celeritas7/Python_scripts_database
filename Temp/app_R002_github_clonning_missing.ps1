# ========= Config =========
$ToolRoot = Split-Path -Parent $MyInvocation.MyCommand.Path   # auto-detect: wherever this script lives
$Python   = "python"                                          # or full path to python.exe if needed

# Folders to hide from the category menu
$ExcludedCategories = @("Powershell", "Old", ".git", "__pycache__")

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
    Write-Host ("-" * $Title.Length) -ForegroundColor DarkGray
    for ($i=0; $i -lt $Items.Count; $i++) {
        "{0,2}) {1}" -f ($i+1), $Items[$i] | Write-Host
    }
    Write-Host ""

    while ($true) {
        $raw = Read-Host "Select 1-$($Items.Count) (or q to quit)"
        if ($raw -eq "q") { return $null }
        $n = 0
        if ([int]::TryParse($raw, [ref]$n) -and $n -ge 1 -and $n -le $Items.Count) {
            return $Items[$n-1]
        }
        Write-Host "Invalid input." -ForegroundColor Red
    }
}

# ========= Step 1: pick category =========
$categories = Get-ChildItem -Path $ToolRoot -Directory |
              Where-Object { $_.Name -notin $ExcludedCategories } |
              Select-Object -ExpandProperty Name |
              Sort-Object

$category = Select-FromList -Title "Choose a category" -Items $categories
if (-not $category) { return }

$catPath = Join-Path $ToolRoot $category

# ========= Step 2: pick task (only root-level scripts, skip Temp/Scrap) =========
$tasks = Get-ChildItem -Path $catPath -File |
         Where-Object { $_.Extension -in ".py", ".ps1" } |
         Select-Object -ExpandProperty Name |
         Sort-Object

$task = Select-FromList -Title "Choose a task in [$category]" -Items $tasks
if (-not $task) { return }

$taskPath = Join-Path $catPath $task

# ========= Confirm & run =========
Write-Host ""
Write-Host "Working folder : $pwd" -ForegroundColor Green
Write-Host "Running        : $task" -ForegroundColor Cyan
Write-Host ""

if ($taskPath.EndsWith(".py")) {
    & $Python $taskPath
}
elseif ($taskPath.EndsWith(".ps1")) {
    & powershell -ExecutionPolicy Bypass -File $taskPath
}
