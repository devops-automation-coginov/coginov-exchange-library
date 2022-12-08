param (
    [Parameter(Mandatory)]
    [string]$ArtefactFolder,
    [Parameter(Mandatory)]
    [string]$base64
)

cd -Path "$ArtefactFolder\Coginov.Exchange.Library\bin\Debug\net6.0\" -PassThru
$fileExec = "$ArtefactFolder\Coginov.Exchange.Library\bin\Debug\net6.0\Coginov.Exchange.Library.dll";
$buffer = [System.Convert]::FromBase64String($base64);
$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($buffer);
Set-AuthenticodeSignature -FilePath $fileExec -Certificate $certificate;