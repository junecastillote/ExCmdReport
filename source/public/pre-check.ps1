# Confirm $PSEdition -eq 'Desktop'
if ($psedition -ne 'Desktop') {
    # Write-Error "This module sucks at PowerShell Core. Use Windows PowerShell 5.1 instead."
    # break
}