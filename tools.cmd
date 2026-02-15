# Create profile if it doesn't exist
if (!(Test-Path $PROFILE)) { New-Item -Path $PROFILE -Force }

# Add the tools function
Add-Content -Path $PROFILE -Value "`nfunction tools { & 'D:\Coding\#Python_scripts_database\app.ps1' }"