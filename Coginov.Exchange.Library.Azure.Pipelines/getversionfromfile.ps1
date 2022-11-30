$version = Get-Content version.txt
Write-Host "##vso[task.setvariable variable=buildNumber]$version"
$Env:buildNumber = "$version"