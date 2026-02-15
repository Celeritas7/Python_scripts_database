# >>> Conda Auto-Initialization >>>
$condaHook = "$env:USERPROFILE\anaconda3\shell\condabin\conda-hook.ps1"
if (Test-Path $condaHook) {
    & $condaHook
    conda activate base
}
# <<< Conda Auto-Initialization <<<

# 🏠 Set default folder when PowerShell starts
Set-Location "C:\Users\manga\OneDrive\#Coding_project_files"

# Import the Chocolatey Profile that contains the necessary code to enable
# tab-completions to function for `choco`.
# Be aware that if you are missing these lines from your profile, tab completion
# for `choco` will not function.
# See https://ch0.co/tab-completion for details.
$ChocolateyProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"
if (Test-Path($ChocolateyProfile)) {
  Import-Module "$ChocolateyProfile"
}

function tools { & 'D:\Coding\#Python_scripts_database\app.ps1' }
