# Get-ChildItem -Path $PSScriptRoot\*.ps1 -Exclude 'InstallMe.ps1','Run.ps1' |
# ForEach-Object {
#     . $_.FullName
# }

$Path = [System.IO.Path]::Combine($PSScriptRoot, 'source')
Get-Childitem $Path -Filter *.ps1 -Recurse | Foreach-Object {
    . $_.Fullname
}