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
$GitCloneOption = "+ Clone a GitHub repo"

$categories = Get-ChildItem -Path $ToolRoot -Directory |
              Where-Object { $_.Name -notin $ExcludedCategories } |
              Select-Object -ExpandProperty Name |
              Sort-Object

# Add clone option at the end
$menuItems = @($categories) + @($GitCloneOption)

$category = Select-FromList -Title "Choose a category" -Items $menuItems
if (-not $category) { return }

# ========= Handle git clone =========
if ($category -eq $GitCloneOption) {
    Write-Host ""
    $repoUrl = Read-Host "Enter GitHub repo URL (e.g. https://github.com/user/repo.git)"
    if ([string]::IsNullOrWhiteSpace($repoUrl)) {
        Write-Host "No URL provided." -ForegroundColor Yellow
        return
    }

    # Extract repo name from URL (strip .git suffix if present)
    $repoName = ($repoUrl -split '/')[-1] -replace '\.git$', ''

    $clonePath = Join-Path $ToolRoot $repoName

    if (Test-Path $clonePath) {
        Write-Host "Folder '$repoName' already exists. Pulling latest..." -ForegroundColor Yellow
        Push-Location $clonePath
        git pull
        Pop-Location
    } else {
        Write-Host "Cloning into: $clonePath" -ForegroundColor Cyan
        git clone $repoUrl $clonePath
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Host "Git operation failed." -ForegroundColor Red
        return
    }

    Write-Host ""
    Write-Host "'$repoName' is now available as a category!" -ForegroundColor Green

    # Set this as the selected category so user can immediately pick a script
    $category = $repoName
}

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
