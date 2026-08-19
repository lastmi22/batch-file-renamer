param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot "dist")
)

$compiler = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$source = Join-Path $PSScriptRoot "src\FileNameBatchRenamer.cs"
$output = Join-Path $OutputDirectory "FileNameBatchRenamer.exe"

if (-not (Test-Path -LiteralPath $compiler)) {
    throw "The .NET Framework 4.x C# compiler was not found: $compiler"
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
& $compiler /nologo /target:winexe "/out:$output" `
    /reference:System.Windows.Forms.dll /reference:System.Drawing.dll $source

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Get-Item -LiteralPath $output
Get-FileHash -LiteralPath $output -Algorithm SHA256
