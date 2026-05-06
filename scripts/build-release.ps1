$ErrorActionPreference = "Stop"

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$releaseDir = Join-Path $root "release"
$nodePath = (Get-Command node).Source
$exePath = Join-Path $releaseDir "organization-scaner.exe"
$seaBlobPath = Join-Path $releaseDir "sea-prep.blob"
$staleNodePath = Join-Path $releaseDir "node.exe"
$postjectPath = Join-Path $root "node_modules\.bin\postject.cmd"
$fuse = "NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2"

New-Item -ItemType Directory -Force $releaseDir | Out-Null

if (Test-Path $staleNodePath) {
  Remove-Item -LiteralPath $staleNodePath -Force
}

Copy-Item -LiteralPath $nodePath -Destination $exePath -Force
Copy-Item -LiteralPath (Join-Path $root "dist") -Destination $releaseDir -Recurse -Force
Copy-Item -LiteralPath (Join-Path $root "node_modules") -Destination $releaseDir -Recurse -Force
Copy-Item -LiteralPath (Join-Path $root "package.json") -Destination $releaseDir -Force
Copy-Item -LiteralPath (Join-Path $root "package-lock.json") -Destination $releaseDir -Force

Push-Location $root
try {
  node --experimental-sea-config scripts\sea-config.json
  & $postjectPath $exePath NODE_SEA_BLOB $seaBlobPath --sentinel-fuse $fuse
  if ($LASTEXITCODE -ne 0) {
    throw "postject failed with exit code $LASTEXITCODE"
  }
}
finally {
  Pop-Location
}

if (Test-Path $seaBlobPath) {
  Remove-Item -LiteralPath $seaBlobPath -Force
}

Write-Host "Release created: $exePath"
