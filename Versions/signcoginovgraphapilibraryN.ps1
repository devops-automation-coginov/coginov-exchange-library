param (
    [Parameter(Mandatory)]
    [string]$ArtefactFolder,
    [Parameter(Mandatory)]
    [string]$SM_CLIENT_CERT_FILE,
    [Parameter(Mandatory)]
    [string]$SM_CLIENT_CERT_PASSWORD,
    [Parameter(Mandatory)]
    [string]$SM_API_KEY,
    [Parameter(Mandatory)]
    [string]$SM_KC_KEY,
    [Parameter(Mandatory)]
    [string]$SM_F_CERT
)

$env:SM_CLIENT_CERT_FILE = $SM_CLIENT_CERT_FILE
$env:SM_CLIENT_CERT_PASSWORD = $SM_CLIENT_CERT_PASSWORD
$env:SM_API_KEY = $SM_API_KEY
$env:SM_HOST="https://clientauth.one.digicert.com"
$path = ""$ArtefactFolder\Coginov.Exchange.Library\bin\Debug\net8.0\"
$pattern = "Coginov*"
$confirmSign = $true
Get-ChildItem -Path $path -Recurse -Include ("$pattern.dll") | ForEach-Object { 
    Write-Host "Archivo encontrado: $($_.FullName)" 
}

# Confirmar si se deben firmar los archivos
if ($confirmSign) {
    Write-Host "Started sign files process..."

    # Procesar cada archivo encontrado
    Get-ChildItem -Path $path -Recurse -Include ("$pattern.dll") | ForEach-Object {
        $filePath = $_.FullName
        Write-Host "Signing file: $filePath"
        
        # Comando para firmar
        & "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe" sign `
            /csp "DigiCert Signing Manager KSP" `
            /kc $SM_KC_KEY `
            /f $SM_F_CERT `
            /tr "http://timestamp.digicert.com" `
            /td SHA256 `
            /fd SHA256 `
            "$filePath"
        
        # Verificar el código de salida del comando de firma
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Error try to sign filed: $filePath. Error code: $LASTEXITCODE"
        } else {
            Write-Host "File singned succesfully: $filePath"
        }
    }
} else {
    Write-Host "Aborted Singning Operation."
}
#END