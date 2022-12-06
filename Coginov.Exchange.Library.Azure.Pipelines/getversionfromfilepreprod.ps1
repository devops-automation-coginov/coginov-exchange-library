$version = Get-Content versionpreprod.txt
Write-Host "##vso[task.setvariable variable=buildNumber]$version"
$Env:buildNumber = "$version"