# ========= Config =========
$ToolRoot = Split-Path -Parent $MyInvocation.MyCommand.Path   # auto-detect: wherever this script lives
$Python   = "python"                                          # or full path to python.exe if needed
$GitUser  = "Celeritas7"                                      # your GitHub username

# Folders to hide from the category menu
$ExcludedCategories = @("Powershell", "Old", ".git", "__pycache__", "Temp", "Scrap")

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

# Replace "Git" folder with built-in Git options
$categories = $categories | Where-Object { $_ -ne "Git" }
$menuItems = @($categories) + @("Git (current folder)")

$category = Select-FromList -Title "Choose a category" -Items $menuItems
if (-not $category) { return }

# ========= Handle built-in Git operations =========
if ($category -eq "Git (current folder)") {
    $gitTasks = @(
        "Pull (git pull origin main)",
        "Sync (add, commit, push)",
        "Clone a repo (by name)",
        "Upload with notes (message + notes, then push)"
    )
    $gitChoice = Select-FromList -Title "Git operation on: $pwd" -Items $gitTasks
    if (-not $gitChoice) { return }

    switch -Wildcard ($gitChoice) {
        "Pull*" {
            Write-Host ""
            Write-Host "Pulling in: $pwd" -ForegroundColor Cyan
            git pull origin main
        }
        "Sync*" {
            Write-Host ""
            if (-not (Test-Path ".git")) {
                Write-Host "ERROR: '$pwd' is not a git repository." -ForegroundColor Red
                return
            }
            Write-Host "Syncing: $pwd" -ForegroundColor Cyan
            git add -A
            $hasChanges = git diff --cached --quiet 2>&1; $changed = $LASTEXITCODE
            if ($changed) {
                git commit -m "Auto sync: update files"
                git push origin main
                Write-Host "Changes pushed!" -ForegroundColor Green
            } else {
                Write-Host "No changes to commit." -ForegroundColor Yellow
            }
        }
        "Clone*" {
            Write-Host ""
            $repoName = Read-Host "Enter repo name (e.g. Cost_management)"
            if ([string]::IsNullOrWhiteSpace($repoName)) {
                Write-Host "No name provided." -ForegroundColor Yellow
                return
            }
            $repoUrl = "https://github.com/$GitUser/$repoName.git"
            $clonePath = Join-Path $pwd $repoName

            if (Test-Path $clonePath) {
                Write-Host "Folder '$repoName' already exists. Pulling latest..." -ForegroundColor Yellow
                Push-Location $clonePath
                git pull origin main
                Pop-Location
            } else {
                Write-Host "Cloning: $repoUrl" -ForegroundColor Cyan
                Write-Host "Into   : $clonePath" -ForegroundColor Cyan
                git clone $repoUrl $clonePath
            }

            if ($LASTEXITCODE -ne 0) {
                Write-Host "Git operation failed." -ForegroundColor Red
            } else {
                Write-Host ""
                Write-Host "'$repoName' cloned successfully!" -ForegroundColor Green
            }
        }
        "Upload*" {
            Write-Host ""

            # --- are we inside a repo? (works from subfolders too) ---
            $inside = git rev-parse --is-inside-work-tree 2>$null
            if ($LASTEXITCODE -ne 0 -or $inside -ne "true") {
                Write-Host "No git repository found in: $pwd" -ForegroundColor Yellow
                $url = Read-Host "Paste GitHub repo URL (Enter to cancel)"
                if ([string]::IsNullOrWhiteSpace($url)) {
                    Write-Host "Cancelled." -ForegroundColor Yellow
                    return
                }
                git init
                if ($LASTEXITCODE -ne 0) {
                    Write-Host "git init failed." -ForegroundColor Red
                    return
                }
                git branch -M main
                git remote add origin $url
                Write-Host "Repository initialised." -ForegroundColor Green
                Write-Host ""
            } else {
                # work from the repo root, not the current subfolder
                $root = git rev-parse --show-toplevel 2>$null
                if ($root) { Set-Location $root }
            }

            # --- make sure an 'origin' remote exists ---
            git remote get-url origin *>$null
            if ($LASTEXITCODE -ne 0) {
                $url = Read-Host "No 'origin' remote. Paste GitHub repo URL (Enter to cancel)"
                if ([string]::IsNullOrWhiteSpace($url)) {
                    Write-Host "Cancelled." -ForegroundColor Yellow
                    return
                }
                git remote add origin $url
            }

            # --- detect branch instead of assuming main ---
            $branch = git branch --show-current 2>$null
            if ([string]::IsNullOrWhiteSpace($branch)) { $branch = "main" }

            Write-Host "Repo   : $pwd"   -ForegroundColor Cyan
            Write-Host "Branch : $branch" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "---------- Changes ----------" -ForegroundColor DarkGray
            git status --short
            Write-Host "-----------------------------" -ForegroundColor DarkGray
            Write-Host ""

            git add -A
            git diff --cached --quiet 2>&1 | Out-Null
            if ($LASTEXITCODE -ne 0) {
                $msg = Read-Host "Commit message (Enter for default)"
                if ([string]::IsNullOrWhiteSpace($msg)) {
                    $msg = "Update $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
                }
                $note = Read-Host "Extra notes (optional, Enter to skip)"

                if ([string]::IsNullOrWhiteSpace($note)) {
                    git commit -m $msg
                } else {
                    git commit -m $msg -m $note
                }

                if ($LASTEXITCODE -ne 0) {
                    Write-Host "Commit failed." -ForegroundColor Red
                    return
                }
            } else {
                Write-Host "No new changes - checking for unpushed commits..." -ForegroundColor Yellow
            }

            # --- pre-push check: only for Node projects with a typecheck/build script ---
            if (Test-Path "package.json") {
                $script = $null
                try {
                    $pkg = Get-Content "package.json" -Raw | ConvertFrom-Json
                    if ($pkg.scripts.PSObject.Properties.Name -contains "typecheck") {
                        $script = "typecheck"
                    } elseif ($pkg.scripts.PSObject.Properties.Name -contains "build") {
                        $script = "build"
                    }
                } catch {
                    Write-Host "Could not read package.json - skipping pre-push check." -ForegroundColor DarkGray
                }

                if ($script) {
                    Write-Host ""
                    Write-Host "Running 'npm run $script' before push (this is what CI runs)..." -ForegroundColor Cyan
                    Write-Host ""
                    npm run $script
                    if ($LASTEXITCODE -ne 0) {
                        Write-Host ""
                        Write-Host "PRE-PUSH CHECK FAILED." -ForegroundColor Red
                        Write-Host "Pushing now will almost certainly fail the deploy workflow." -ForegroundColor Yellow
                        $go = Read-Host "Push anyway? (y/N)"
                        if ($go -ne "y") {
                            Write-Host "Push cancelled. Commit is saved locally - fix the errors and run this again." -ForegroundColor Yellow
                            return
                        }
                    } else {
                        Write-Host ""
                        Write-Host "Pre-push check passed." -ForegroundColor Green
                    }
                }
            }

            Write-Host ""
            Write-Host "Pushing to origin/$branch ..." -ForegroundColor Cyan
            git push -u origin $branch

            if ($LASTEXITCODE -ne 0) {
                Write-Host ""
                Write-Host "PUSH FAILED - the remote likely has commits you don't have." -ForegroundColor Red
                Write-Host "Run the Pull option first, then try again." -ForegroundColor Yellow
            } else {
                Write-Host ""
                Write-Host "Done - files uploaded to GitHub." -ForegroundColor Green
            }
        }
    }
    return
}

$catPath = Join-Path $ToolRoot $category

# ========= Step 2: pick task (only root-level scripts, skip Temp/Scrap) =========
$tasks = Get-ChildItem -Path $catPath -File |
         Where-Object { $_.Extension -in ".py", ".ps1", ".bat" } |
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
elseif ($taskPath.EndsWith(".bat")) {
    & cmd /c $taskPath
}