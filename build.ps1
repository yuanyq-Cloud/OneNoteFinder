<#
.SYNOPSIS
    OneFinder 一键构建 + 打包 MSI

.DESCRIPTION
    步骤：
      1. dotnet publish  — 将 .NET 8 应用发布为 win-x64 自包含单文件 EXE
      2. wix build       — 将 EXE 打包成 MSI 安装包

.PREREQUISITES
    - .NET 8 SDK      : winget install Microsoft.DotNet.SDK.8
    - WiX Toolset v5  : dotnet tool install --global wix

.EXAMPLE
    .\build.ps1                   # Release 配置
    .\build.ps1 -Configuration Debug
#>

param(
    [string] $Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$Root           = $PSScriptRoot
$AppProject     = Join-Path $Root "OneFinder\OneFinder.csproj"
$WxsFile        = Join-Path $Root "installer\Package.wxs"
$PublishDir     = Join-Path $Root "OneFinder\bin\$Configuration\net8.0-windows\win-x64\publish\"
$OutputMsi      = Join-Path $Root "installer\OneFinderSetup.msi"

# ── 验证前置工具 ──────────────────────────────────────────────
function Require-Command([string]$name) {
    if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
        Write-Error "未找到命令 '$name'。请先安装所需工具（见脚本注释）。"
        exit 1
    }
}

Require-Command "dotnet"
Require-Command "wix"

# ── 步骤 1：发布自包含单文件 EXE ─────────────────────────────
Write-Host ""
Write-Host ">>> [1/2] 发布 OneFinder (self-contained, win-x64, single-file)..." -ForegroundColor Cyan

dotnet publish $AppProject `
    --configuration $Configuration `
    --runtime win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:PublishTrimmed=false `
    -p:DebugType=none `
    -p:DebugSymbols=false

if ($LASTEXITCODE -ne 0) { Write-Error "dotnet publish 失败"; exit 1 }

# 确认 EXE 存在
$ExePath = Join-Path $PublishDir "OneFinder.exe"
if (-not (Test-Path $ExePath)) {
    Write-Error "未找到发布产物：$ExePath"
    exit 1
}
Write-Host "    EXE: $ExePath" -ForegroundColor Gray

# ── 步骤 2：WiX 构建 MSI ───────────────────────────────────────
Write-Host ""
Write-Host ">>> [2/2] 打包 MSI..." -ForegroundColor Cyan

# 确保 PublishDir 结尾有反斜杠（WiX 文件路径拼接需要）
if (-not $PublishDir.EndsWith("\")) { $PublishDir += "\" }

wix build $WxsFile `
    -d "PublishDir=$PublishDir" `
    -arch x64 `
    -out $OutputMsi

if ($LASTEXITCODE -ne 0) { Write-Error "wix build 失败"; exit 1 }

Write-Host ""
Write-Host ">>> 完成！MSI 位置：" -ForegroundColor Green
Write-Host "    $OutputMsi" -ForegroundColor Yellow
