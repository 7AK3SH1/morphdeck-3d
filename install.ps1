param(
    [ValidateSet("antigravity", "claude", "codex", "cursor", "all")]
    [string]$Target = "antigravity",
    [string]$WorkspaceDir = "."
)

$ErrorActionPreference = "Stop"
$RepoRoot = $PSScriptRoot

$TargetMap = @{
    "antigravity" = Join-Path $env:USERPROFILE ".gemini\config\skills\morphdeck-3d"
    "claude"      = Join-Path $env:USERPROFILE ".claude\skills\morphdeck-3d"
    "codex"       = Join-Path $env:USERPROFILE ".codex\skills\morphdeck-3d"
    "cursor"      = Join-Path (Resolve-Path $WorkspaceDir) ".cursor\skills\morphdeck-3d"
}

$TargetsToInstall = if ($Target -eq "all") {
    @("antigravity", "claude", "codex", "cursor")
} else {
    @($Target)
}

$CopyItems = @("SKILL.md", "LICENSE", "NOTICE", "requirements.txt", "scripts", "references", "examples")

foreach ($t in $TargetsToInstall) {
    $dest = $TargetMap[$t]
    if (Test-Path $dest) {
        Remove-Item -Path $dest -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $dest | Out-Null

    foreach ($item in $CopyItems) {
        $src = Join-Path $RepoRoot $item
        if (Test-Path $src) {
            Copy-Item -Path $src -Destination (Join-Path $dest $item) -Recurse -Force
        }
    }

    Get-ChildItem -Path $dest -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    Write-Host "Installed morphdeck-3d ($t) -> $dest"
}
