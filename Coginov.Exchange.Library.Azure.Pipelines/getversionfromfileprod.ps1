$version = Get-Content versionprod.txt
Write-Host "##vso[task.setvariable variable=buildNumber]$version"
$Env:buildNumber = "$version"