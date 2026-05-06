$ErrorActionPreference = "Stop"

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$releaseDir = Join-Path $root "release"
$nodePath = (Get-Command node).Source
$exePath = Join-Path $releaseDir "ya-finder.exe"
$seaBlobPath = Join-Path $releaseDir "sea-prep.blob"
$bundlePath = Join-Path $releaseDir "app.cjs"
$postjectPath = Join-Path $root "node_modules\.bin\postject.cmd"
$esbuildPath = Join-Path $root "node_modules\.bin\esbuild.cmd"
$fuse = "NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2"

if (Test-Path $releaseDir) {
  Remove-Item -LiteralPath $releaseDir -Recurse -Force
}

New-Item -ItemType Directory -Force $releaseDir | Out-Null

& $esbuildPath (Join-Path $root "dist\index.js") `
  --bundle `
  --platform=node `
  --format=cjs `
  --target=node22 `
  --outfile=$bundlePath `
  "--banner:js=process.emitWarning=function(){};" `
  --packages=bundle `
  "--alias:supports-color=./scripts/stubs/supports-color.cjs" `
  "--alias:utf-8-validate=./scripts/stubs/utf-8-validate.cjs" `
  "--alias:bufferutil=./scripts/stubs/bufferutil.cjs"
if ($LASTEXITCODE -ne 0) {
  throw "esbuild failed with exit code $LASTEXITCODE"
}

Copy-Item -LiteralPath $nodePath -Destination $exePath -Force

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

if (Test-Path $bundlePath) {
  Remove-Item -LiteralPath $bundlePath -Force
}

Write-Host "Release created: $exePath"
