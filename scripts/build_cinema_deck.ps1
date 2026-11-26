# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
# Master Cinema-Morph v4.0.0 PowerPoint COM Builder for morphdeck-3d

param(
    [string]$SpecJson = "",
    [string]$AssetsDir = "",
    [Parameter(Mandatory = $true)][string]$OutputPptx,
    [Parameter(Mandatory = $true)][string]$ExportDir,
    [string]$CopyTo = "",
    [switch]$SkipExport,
    [switch]$ExportOnly
)

$ErrorActionPreference = "Stop"

if ($ExportOnly) {
    if (!(Test-Path $OutputPptx)) { throw "OutputPptx not found for -ExportOnly: $OutputPptx" }
    if (!(Test-Path $ExportDir)) { New-Item -ItemType Directory -Force -Path $ExportDir | Out-Null }
    Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $null
    try {
        $pres = $ppt.Presentations.Open($OutputPptx, -1, 0, 0)
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $pngPath = Join-Path $ExportDir "slide_$i.png"
            $pres.Slides.Item($i).Export($pngPath, "PNG", 1920, 1080)
            Write-Host "Exported slide $i -> $pngPath"
        }
        if (-not [string]::IsNullOrWhiteSpace($CopyTo)) {
            Copy-Item -Path $OutputPptx -Destination $CopyTo -Force
            Write-Host ("Copied deck -> " + $CopyTo)
        }
    } finally {
        if ($null -ne $pres) { try { $pres.Close() } catch {} }
        try { $ppt.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    exit 0
}

$resolvedPath = Join-Path $AssetsDir "deck_resolved.json"
if (Test-Path $resolvedPath) {
    $spec = Get-Content -Raw -Encoding UTF8 $resolvedPath | ConvertFrom-Json
} else {
    $spec = Get-Content -Raw -Encoding UTF8 $SpecJson | ConvertFrom-Json
}

$fontsInstalled = (Get-ChildItem "$env:WINDIR\Fonts" -ErrorAction SilentlyContinue).Name
function Resolve-InstalledFont([string]$preferred, [string]$fallback = "Segoe UI") {
    $p = if ($preferred) { $preferred } else { $fallback }
    if ($p -eq "Georgia" -and ($fontsInstalled -contains "georgia.ttf")) { return "Georgia" }
    if ($p -eq "Impact" -and ($fontsInstalled -contains "impact.ttf")) { return "Impact" }
    if ($p -eq "Consolas" -and ($fontsInstalled -contains "consola.ttf")) { return "Consolas" }
    if ($p -eq "Bahnschrift" -and ($fontsInstalled -contains "bahnschrift.ttf")) { return "Bahnschrift" }
    if ($p -like "Segoe UI*" -and ($fontsInstalled -contains "segoeui.ttf")) { return $p }
    if ($fontsInstalled -contains "bahnschrift.ttf") { return "Bahnschrift" }
    return "Arial"
}

$DisplayFont = Resolve-InstalledFont $spec.fonts.display "Bahnschrift"
$HeadingFont = Resolve-InstalledFont $spec.fonts.heading $DisplayFont
$BodyFont    = Resolve-InstalledFont $spec.fonts.body "Segoe UI"
$SemiFont    = Resolve-InstalledFont $spec.fonts.semi "Segoe UI Semibold"
$MonoFont    = Resolve-InstalledFont $spec.fonts.mono "Consolas"

function HexToBgr([string]$hex, [string]$fallback = "#2563EB") {
    $h = if ([string]::IsNullOrWhiteSpace($hex)) { $fallback } else { $hex }
    $h = $h.Trim().TrimStart("#")
    if ($h.Length -ne 6) { $h = $fallback.TrimStart("#") }
    $r = [Convert]::ToInt32($h.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($h.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
    return [int]($r + ($g * 256) + ($b * 65536))
}

function Find-Shape($slide, [string]$name) {
    if ($null -eq $slide) { return $null }
    for ($k = 1; $k -le $slide.Shapes.Count; $k++) {
        $sh = $slide.Shapes.Item($k)
        if ($sh.Name -eq $name) { return $sh }
    }
    return $null
}

function Clean-SlideExcept($slide, [string[]]$keepExactNames) {
    for ($k = $slide.Shapes.Count; $k -ge 1; $k--) {
        if ($keepExactNames -notcontains $slide.Shapes.Item($k).Name) {
            $slide.Shapes.Item($k).Delete()
        }
    }
}

function Start-MorphSlide($pres, $prevSlide, [string[]]$keepExactNames) {
    if ($null -ne $prevSlide -and $keepExactNames.Count -gt 0) {
        $s = ($prevSlide.Duplicate()).Item(1)
        Clean-SlideExcept $s $keepExactNames
        return $s
    }
    $s = $pres.Slides.Add($pres.Slides.Count + 1, 12)
    $s.FollowMasterBackground = 0
    return $s
}

function Set-Morph($slide, [single]$dur) {
    $slide.SlideShowTransition.EntryEffect = 3954 # ppEffectMorph
    $slide.SlideShowTransition.Duration = $dur
    $slide.SlideShowTransition.Speed = 1
}

function Add-Pic($slide, [string]$name, [string]$imgPath, [single]$x, [single]$y, [single]$w, [single]$h) {
    if ([string]::IsNullOrWhiteSpace($imgPath) -or !(Test-Path $imgPath)) { return $null }
    $existing = Find-Shape $slide $name
    if ($null -ne $existing) {
        try { $existing.PictureFormat.CropTop = [single]0; $existing.PictureFormat.CropBottom = [single]0 } catch {}
        $existing.AlternativeText = ""
        $existing.Left = $x; $existing.Top = $y; $existing.Width = $w; $existing.Height = $h
        $existing.Rotation = [single]0
        return $existing
    }
    $sh = $slide.Shapes.AddPicture($imgPath, 0, 1, $x, $y, $w, $h)
    $sh.Name = $name
    return $sh
}

function Add-Rect($slide, [string]$name, [single]$x, [single]$y, [single]$w, [single]$h, [int]$fillBgr, [single]$fillTrans = 0, [int]$borderBgr = -1, [single]$borderWeight = 0, [single]$cornerAdj = 0.0) {
    $existing = Find-Shape $slide $name
    if ($null -ne $existing) {
        $existing.Left = $x; $existing.Top = $y; $existing.Width = $w; $existing.Height = $h
        if ($fillTrans -ge 1.0) { $existing.Fill.Visible = 0 }
        else { $existing.Fill.Visible = -1; $existing.Fill.Solid(); $existing.Fill.ForeColor.RGB = $fillBgr; $existing.Fill.Transparency = $fillTrans }
        if ($borderBgr -ge 0 -and $borderWeight -gt 0) {
            $existing.Line.Visible = -1; $existing.Line.ForeColor.RGB = $borderBgr; $existing.Line.Weight = $borderWeight
        } else { $existing.Line.Visible = 0 }
        return $existing
    }
    $stype = if ($cornerAdj -gt 0.0) { 5 } else { 1 }
    $sh = $slide.Shapes.AddShape($stype, $x, $y, $w, $h)
    $sh.Name = $name
    if ($cornerAdj -gt 0.0) { $sh.Adjustments.Item(1) = $cornerAdj }
    if ($fillTrans -ge 1.0) { $sh.Fill.Visible = 0 }
    else { $sh.Fill.Solid(); $sh.Fill.ForeColor.RGB = $fillBgr; $sh.Fill.Transparency = $fillTrans }
    if ($borderBgr -ge 0 -and $borderWeight -gt 0) {
        $sh.Line.Visible = -1; $sh.Line.ForeColor.RGB = $borderBgr; $sh.Line.Weight = $borderWeight
    } else { $sh.Line.Visible = 0 }
    return $sh
}

function Add-Text($slide, [string]$name, [string]$text, [single]$x, [single]$y, [single]$w, [single]$h, [string]$font, [single]$size, [bool]$bold, [int]$bgr, [int]$align = 1, [int]$vAnchor = 1) {
    $tb = Find-Shape $slide $name
    if ($null -eq $tb) {
        $tb = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
        $tb.Name = $name
    } else {
        $tb.Left = $x; $tb.Top = $y; $tb.Width = $w; $tb.Height = $h
    }
    $tf = $tb.TextFrame2
    $tf.WordWrap = -1; $tf.AutoSize = 0
    $tf.MarginLeft = [single]2; $tf.MarginRight = [single]2; $tf.MarginTop = [single]2; $tf.MarginBottom = [single]2
    $tf.VerticalAnchor = $vAnchor
    $tr = $tf.TextRange
    $tr.Text = if ($null -eq $text) { "" } else { [string]$text }
    $tr.Font.Name = $font; $tr.Font.Size = $size; $tr.Font.Bold = if ($bold) { -1 } else { 0 }
    $tr.Font.Fill.Solid(); $tr.Font.Fill.ForeColor.RGB = $bgr
    $tr.ParagraphFormat.Alignment = $align
    return $tb
}

function Add-TextShadow($tb) {
    try {
        $sh = $tb.TextFrame2.TextRange.Font.Shadow
        $sh.Visible = -1; $sh.Blur = [single]8; $sh.OffsetX = [single]0; $sh.OffsetY = [single]2.5; $sh.Transparency = [single]0.44
    } catch {}
}

function Add-CapsulePill($slide, [string]$name, [string]$text, [single]$x, [single]$y, [single]$w, [single]$h, [int]$bgBgr, [single]$bgTrans, [int]$fgBgr, [single]$fontSize, [string]$fontName = $MonoFont) {
    $sh = Find-Shape $slide $name
    if ($null -eq $sh) {
        $sh = $slide.Shapes.AddShape(5, $x, $y, $w, $h)
        $sh.Name = $name
        $sh.Adjustments.Item(1) = [single]0.5
    } else { $sh.Left = $x; $sh.Top = $y; $sh.Width = $w; $sh.Height = $h }
    $sh.Fill.Solid(); $sh.Fill.ForeColor.RGB = $bgBgr; $sh.Fill.Transparency = $bgTrans; $sh.Line.Visible = 0
    $tf = $sh.TextFrame2; $tf.WordWrap = 0; $tf.AutoSize = 0; $tf.VerticalAnchor = 3
    $tf.MarginLeft = [single]8; $tf.MarginRight = [single]8
    $tr = $tf.TextRange; $tr.Text = [string]$text
    $tr.Font.Name = $fontName; $tr.Font.Size = $fontSize; $tr.Font.Bold = -1
    $tr.Font.Fill.Solid(); $tr.Font.Fill.ForeColor.RGB = $fgBgr; $tr.ParagraphFormat.Alignment = 2
    return $sh
}

function Ensure-Model3D($slide, [string]$name, [string]$glbPath, [single]$x, [single]$y, [single]$w, [single]$h, [single]$rx, [single]$ry, [single]$rz) {
    if ([string]::IsNullOrWhiteSpace($glbPath) -or !(Test-Path $glbPath)) { return $null }
    $m = Find-Shape $slide $name
    if ($null -eq $m) {
        $m = $slide.Shapes.Add3DModel($glbPath, 0, 1, $x, $y, $w, $h)
        $m.Name = $name
        $m.LockAspectRatio = -1
        $m.Model3D.AutoFit = 0
    } else {
        $m.Left = $x; $m.Top = $y; $m.Width = $w; $m.Height = $h
    }
    $m.Model3D.RotationX = $rx; $m.Model3D.RotationY = $ry; $m.Model3D.RotationZ = $rz
    $m.ZOrder(0)
    return $m
}

function Add-OdometerStrip($slide, [string]$name, [single]$x, [single]$y, [single]$w, [int]$txtBgr, [string]$font) {
    $tb = Find-Shape $slide $name
    if ($null -eq $tb) {
        $tb = $slide.Shapes.AddTextbox(1, $x, $y, $w, [single]920)
        $tb.Name = $name
    } else {
        $tb.Left = $x; $tb.Top = $y; $tb.Width = $w; $tb.Height = [single]920
    }
    $tf = $tb.TextFrame2
    $tf.WordWrap = 0; $tf.AutoSize = 0; $tf.VerticalAnchor = 1
    $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $tr = $tf.TextRange
    $tr.Text = "0`r1`r2`r3`r4`r5`r6`r7`r8`r9"
    $tr.Font.Name = $font; $tr.Font.Size = [single]66; $tr.Font.Bold = -1
    $tr.Font.Fill.Solid(); $tr.Font.Fill.ForeColor.RGB = $txtBgr
    $pf = $tr.ParagraphFormat
    $pf.Alignment = 2
    $pf.LineRuleWithin = 0
    $pf.SpaceWithin = [single]86
    $pf.SpaceBefore = 0
    $pf.SpaceAfter = 0
    return $tb
}

$isLight   = [bool]$spec.is_light
$cBg       = HexToBgr $spec.palette.bg "#080A0F"
$cSurface  = HexToBgr $spec.palette.surface "#141824"
$cSurfAlt  = HexToBgr $spec.palette.surface_alt "#1C2233"
$cInk      = HexToBgr $spec.palette.ink "#F8FAFC"
$cMuted    = HexToBgr $spec.palette.ink_muted "#94A3B8"
$cBorder   = HexToBgr $spec.palette.border "#334155"
$cAcc1     = HexToBgr $spec.palette.accent1 "#E11D48"
$cAcc2     = HexToBgr $spec.palette.accent2 "#38BDF8"
$cAcc3     = HexToBgr $spec.palette.accent3 "#F59E0B"
$cWhite    = HexToBgr "#FFFFFF"
$cObsidian = HexToBgr "#080A0F"

$heroGlb = if ($spec.resolved_glb) { [string]$spec.resolved_glb } else { "" }

if (!(Test-Path $ExportDir)) { New-Item -ItemType Directory -Force -Path $ExportDir | Out-Null }
$outParent = Split-Path -Parent $OutputPptx
if ($outParent -and !(Test-Path $outParent)) { New-Item -ItemType Directory -Force -Path $outParent | Out-Null }

Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $null

try {
    $pres = $ppt.Presentations.Add(0)
    $pres.PageSetup.SlideWidth = [single]960
    $pres.PageSetup.SlideHeight = [single]540

    $slides = @($spec.slides)
    $prevSlide = $null
    $prevType  = ""
    $prevSl    = $null

    for ($si = 0; $si -lt $slides.Count; $si++) {
        $sl = $slides[$si]
        $idx = $si + 1
        $st = [string]$sl.type
        $nextSl = if (($si + 1) -lt $slides.Count) { $slides[$si + 1] } else { $null }
        $nextType = if ($null -ne $nextSl) { [string]$nextSl.type } else { "" }
        Write-Host ("Building Slide {0}/{1} ({2} | universe={3})..." -f $idx, $slides.Count, $st, $spec.universe)

        # Compute incoming pre-staged shapes we MUST carry over from $prevSlide
        $carryBridge = @()
        if ($prevType -eq "real_portal_flythrough" -and $st -ne "real_portal_flythrough") {
            $carryBridge += @("!!portal_bg", "!!portal_inner")
        }
        if ($prevType -eq "vertical_world_elevator" -and $st -ne "vertical_world_elevator") {
            $carryBridge += @("!!elev_bg", "!!elev_fg")
        }

        switch ($st) {
            "newspaper_dive" {
                $stage = if ($sl.stage) { [string]$sl.stage } else { "frontpage" }
                $keep = $carryBridge
                if ($prevType -eq "newspaper_dive") { $keep += @("!!news_sheet", "!!ns_cloud", "!!ns_h1") }
                $s = Start-MorphSlide $pres $prevSlide $keep

                if ($stage -eq "collage") {
                    $ns = Add-Pic $s "!!news_sheet" "$AssetsDir\news_sheet.png" -1180 -680 3456 1944; $ns.ZOrder(1)
                    Add-Pic $s "!!col_cutout_L" $sl._left_cutout 32 94 336 446 | Out-Null
                    Add-Pic $s "!!col_cutout_R" $sl._right_cutout 468 156 410 384 | Out-Null
                    Add-Pic $s "!!ns_cloud" "$AssetsDir\ep_paper_cloud.png" 476 30 448 224 | Out-Null
                    Add-Text $s "!!ns_ovr" $sl.overline 514 48 376 18 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "!!ns_h1" $sl.headline 514 68 376 66 $DisplayFont 23 $true $cAcc1 1 1 | Out-Null
                    Add-Text $s "!!ns_body" $sl.body 514 140 376 82 $BodyFont 11.5 $false $cInk 1 1 | Out-Null
                    Set-Morph $s ([single]2.0)
                } elseif ($stage -eq "masthead_spread") {
                    $ns = Add-Pic $s "!!news_sheet" "$AssetsDir\news_sheet.png" -680 -320 2400 1350; $ns.ZOrder(1)
                    $nc = Add-Pic $s "!!ns_cloud" "$AssetsDir\ep_paper_cloud.png" 160 148 640 260; $nc.ZOrder(0)
                    Add-Text $s "!!ns_ovr" $sl.overline 210 182 540 22 $MonoFont 9.5 $true $cAcc1 2 3 | Out-Null
                    Add-Text $s "!!ns_h1" $sl.headline 200 210 560 84 $DisplayFont 30 $true $cInk 2 3 | Out-Null
                    Add-Text $s "!!ns_body" $sl.body 220 300 520 76 $BodyFont 12.5 $false $cMuted 2 1 | Out-Null
                    Set-Morph $s ([single]1.9)
                } else {
                    $ns = Add-Pic $s "!!news_sheet" "$AssetsDir\news_sheet.png" 0 0 960 540; $ns.ZOrder(1)
                    Add-Pic $s "!!ns_cloud" "$AssetsDir\ep_paper_cloud.png" 160 620 640 260 | Out-Null
                    Add-Text $s "!!ns_h1" $sl.headline 200 660 560 84 $DisplayFont 30 $true $cInk 2 3 | Out-Null
                    Set-Morph $s ([single]1.6)
                }
            }

            "editorial_print_overlay" {
                $keep = $carryBridge + @("!!news_sheet", "!!ep_drips", "!!ep_port", "!!ep_cloud")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $ns = Add-Pic $s "!!news_sheet" "$AssetsDir\news_sheet.png" -280 -180 1440 810; $ns.ZOrder(1)
                Add-Rect $s "ep_scrim" 0 0 960 540 $cBg ([single]0.18) -1 0 0 | Out-Null

                $dr = Add-Pic $s "!!ep_drips" "$AssetsDir\ep_drips.png" 0 0 960 120; $dr.ZOrder(0)
                $portSrc = if ($sl._portrait_cutout) { [string]$sl._portrait_cutout } elseif (Test-Path "$AssetsDir\hero_cutout_auto.png") { "$AssetsDir\hero_cutout_auto.png" } else { "" }
                if ($portSrc) { $ep = Add-Pic $s "!!ep_port" $portSrc 48 86 344 444; $ep.ZOrder(0) }

                $ec = Add-Pic $s "!!ep_cloud" "$AssetsDir\ep_paper_cloud.png" 396 118 528 356; $ec.ZOrder(0)
                Add-Text $s "!!ep_ovr" $sl.overline 446 162 430 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "!!ep_h1" $sl.headline 446 190 430 74 $DisplayFont 26 $true $cAcc1 1 1 | Out-Null
                Add-Text $s "!!ep_body" $sl.body 446 272 430 138 $BodyFont 12.5 $false $cInk 1 1 | Out-Null
                Set-Morph $s ([single]1.9)
            }

            "odometer_counter" {
                $keep = $carryBridge + @("!!odo_curtain", "!!odo_mask_top", "!!odo_mask_bot", "!!odo_sil")
                $targetStr = [string]$sl.target_value
                $chars = $targetStr.ToCharArray()
                for ($ci = 0; $ci -lt $chars.Count; $ci++) {
                    if ($chars[$ci] -match '\d') { $keep += "!!odo_col_$ci" }
                }
                $s = Start-MorphSlide $pres $prevSlide $keep
                $odoBg = if ($sl.bg_hex) { HexToBgr $sl.bg_hex } else { $cObsidian }
                $curt = Add-Rect $s "!!odo_curtain" 0 0 960 540 $odoBg 0 -1 0 0; $curt.ZOrder(1)

                $winY = [single]218; $stepH = [single]86
                $totW = 0
                foreach ($ch in $chars) { $totW += if ($ch -eq ' ') { 24 } elseif ($ch -match '\d') { 50 } else { 32 } }
                $curX = [single](480 - ($totW / 2))

                for ($ci = 0; $ci -lt $chars.Count; $ci++) {
                    $ch = $chars[$ci]
                    if ($ch -eq ' ') { $curX += 24; continue }
                    if ($ch -match '\d') {
                        $d = [int][string]$ch
                        $yTarget = [single]($winY - ($d * $stepH))
                        $st_sh = Add-OdometerStrip $s "!!odo_col_$ci" $curX $yTarget 50 $cWhite $HeadingFont
                        $st_sh.ZOrder(0)
                        $curX += 50
                    } else {
                        $sep = Add-Text $s "odo_sep_$ci" ([string]$ch) $curX $winY 32 86 $HeadingFont 66 $true $cWhite 2 3
                        $sep.ZOrder(0)
                        $curX += 32
                    }
                }

                $mTop = Add-Rect $s "!!odo_mask_top" -4 -4 968 222 $odoBg 0 -1 0 0; $mTop.ZOrder(0)
                $mBot = Add-Rect $s "!!odo_mask_bot" -4 304 968 242 $odoBg 0 -1 0 0; $mBot.ZOrder(0)

                $silPath = if ($sl._silhouette_cutout) { [string]$sl._silhouette_cutout } else { "$AssetsDir\odo_silhouette.png" }
                $sil = Add-Pic $s "!!odo_sil" $silPath 0 378 960 162; $sil.ZOrder(0)
                $ovr = Add-Text $s "!!odo_ovr" $sl.overline 120 146 720 32 $DisplayFont 20 $true $cAcc1 2 3; $ovr.ZOrder(0)
                $sub = Add-Text $s "!!odo_sub" $sl.sublabel 120 322 720 34 $DisplayFont 22 $true $cWhite 2 3; $sub.ZOrder(0)
                Set-Morph $s ([single]2.1)
            }

            "flyby_wipe" {
                $keep = $carryBridge + @("!!flyby_obj", "!!flyby_cloud")
                if ($heroGlb -and ($sl.use_hero_3d -ne $false)) { $keep += @("!!HeroShadow", "!!Hero3D") }
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bgP = if ($sl._bg_image) { [string]$sl._bg_image } else { "$AssetsDir\bg_base.png" }
                $bg = Add-Pic $s "!!StageBg" $bgP 0 0 960 540; $bg.ZOrder(1)

                $fc = Add-Pic $s "!!flyby_cloud" "$AssetsDir\flyby_cloud.png" 980 -220 760 440; $fc.ZOrder(0)
                if ($sl._flyby_cutout) {
                    $fo = Add-Pic $s "!!flyby_obj" $sl._flyby_cutout 1140 -300 440 280
                    $fo.Rotation = [single]-14; $fo.ZOrder(0)
                }

                $txtCol = if ($sl._bg_image) { $cWhite } else { $cInk }
                Add-Text $s "!!fb_ovr" $sl.overline 56 56 480 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                $fh1 = Add-Text $s "!!fb_h1" $sl.headline 56 84 490 92 $DisplayFont 34 $true $txtCol 1 1; if ($sl._bg_image) { Add-TextShadow $fh1 }
                Add-Text $s "!!fb_body" $sl.body 56 188 460 108 $BodyFont 13 $false $txtCol 1 1 | Out-Null

                if ($heroGlb -and ($sl.use_hero_3d -ne $false)) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 552 430 352 54; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 560 92 336 360 12 -24 2 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "cutout_map_stage" {
                $keep = $carryBridge + @("!!cm_map", "!!cm_div", "!!cm_h1")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bgP = if ($sl._bg_image) { [string]$sl._bg_image } else { "$AssetsDir\bg_base.png" }
                $bg = Add-Pic $s "!!StageBg" $bgP 0 0 960 540; $bg.ZOrder(1)

                if ($sl._map_image) {
                    $mp = Add-Pic $s "!!cm_map" $sl._map_image 42 54 520 432; $mp.ZOrder(0)
                }
                Add-Rect $s "!!cm_div" 588 68 ([single]1.5) 404 $cBorder 0 -1 0 0 | Out-Null
                Add-Text $s "!!cm_ovr" $sl.overline 614 72 294 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "!!cm_h1" $sl.headline 614 100 294 84 $DisplayFont 30 $true $cAcc1 1 1 | Out-Null
                Add-Text $s "!!cm_body" $sl.body 614 196 294 150 $BodyFont 12.5 $false $cInk 1 1 | Out-Null
                $kpis = @($sl.kpis)
                for ($ki = 0; $ki -lt [Math]::Min(2, $kpis.Count); $ki++) {
                    $kx = [single](614 + $ki * 150)
                    Add-Rect $s "!!cm_kline_$ki" $kx 368 132 2 $cAcc1 0 -1 0 0 | Out-Null
                    Add-Text $s "!!cm_kv_$ki" $kpis[$ki].value $kx 378 132 40 $DisplayFont 26 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "!!cm_kl_$ki" $kpis[$ki].label $kx 420 132 22 $MonoFont 9 $false $cMuted 1 1 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "poster_depth_reveal" {
                $keep = $carryBridge + @("!!poster_band", "!!poster_word", "!!poster_hero", "!!HeroShadow", "!!Hero3D")
                $s = Start-MorphSlide $pres $prevSlide $keep

                if ($nextType -eq "picture_frame_pullback") {
                    $wallNext = if ($nextSl._wall_image) { [string]$nextSl._wall_image } else { "$AssetsDir\bg_alt.png" }
                    $wr = Add-Pic $s "!!wall_room" $wallNext -1632 -918 4224 2376; $wr.ZOrder(1)
                }

                $bgP = if ($sl._poster_bg_image) { [string]$sl._poster_bg_image } elseif ($sl._bg_image) { [string]$sl._bg_image } else { "$AssetsDir\bg_alt.png" }
                $pbg = Add-Pic $s "!!frame_art" $bgP 0 0 960 540; $pbg.ZOrder(0)

                $band = Add-Rect $s "!!poster_band" 348 0 264 540 $cAcc1 0 -1 0 0; $band.ZOrder(0)
                $gw = if ($sl.giant_word) { [string]$sl.giant_word } else { "2026" }
                # Giant word cleanly in upper third; 3D hero below so zero collision!
                $pw = Add-Text $s "!!poster_word" $gw 60 24 840 138 $HeadingFont 116 $true $cWhite 2 3; Add-TextShadow $pw; $pw.ZOrder(0)

                if ($sl._hero_cutout) {
                    $ph = Add-Pic $s "!!poster_hero" $sl._hero_cutout 284 152 392 368; $ph.ZOrder(0)
                } elseif ($heroGlb -and ($sl.use_hero_3d -ne $false)) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 296 474 368 54; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 296 152 368 364 12 -22 2 | Out-Null
                }

                if ($nextType -eq "picture_frame_pullback" -and $nextSl._hollow_frame) {
                    $hf = Add-Pic $s "!!hollow_frame" $nextSl._hollow_frame -240 -180 1440 900; $hf.ZOrder(0)
                }
                Set-Morph $s ([single]2.0)
            }

            "picture_frame_pullback" {
                # Keep !!Hero3D alive so the live 3D model shrinks concentrically straight into the museum frame!
                $keep = $carryBridge + @("!!wall_room", "!!frame_art", "!!hollow_frame", "!!poster_band", "!!poster_word", "!!poster_hero", "!!Hero3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $wallP = if ($sl._wall_image) { [string]$sl._wall_image } else { "$AssetsDir\bg_alt.png" }
                $wbg = Add-Pic $s "!!wall_room" $wallP 0 0 960 540; $wbg.ZOrder(1)

                $fx = if ($null -ne $sl.frame_x) { [single]$sl.frame_x } else { [single]508 }
                $fy = if ($null -ne $sl.frame_y) { [single]$sl.frame_y } else { [single]78 }
                $fw = if ($null -ne $sl.frame_w) { [single]$sl.frame_w } else { [single]340 }
                $fh = if ($null -ne $sl.frame_h) { [single]$sl.frame_h } else { [single]192 }

                $artSrc = if ($sl._framed_art) { [string]$sl._framed_art } elseif ($null -ne (Find-Shape $s "!!frame_art")) { "" } else { "$AssetsDir\bg_alt.png" }
                if ($artSrc) {
                    $fa = Add-Pic $s "!!frame_art" $artSrc $fx $fy $fw $fh; $fa.ZOrder(0)
                } elseif ($null -ne (Find-Shape $s "!!frame_art")) {
                    $fa = Find-Shape $s "!!frame_art"
                    $fa.Left = $fx; $fa.Top = $fy; $fa.Width = $fw; $fa.Height = $fh; $fa.ZOrder(0)
                }

                $pb = Find-Shape $s "!!poster_band"
                if ($null -ne $pb) {
                    $pb.Left = [single]($fx + $fw * 0.362); $pb.Top = $fy
                    $pb.Width = [single]($fw * 0.276); $pb.Height = $fh; $pb.ZOrder(0)
                }
                $pw = Find-Shape $s "!!poster_word"
                if ($null -ne $pw) {
                    $pw.Left = [single]($fx + $fw * 0.08); $pw.Top = [single]($fy + $fh * 0.04)
                    $pw.Width = [single]($fw * 0.84); $pw.Height = [single]($fh * 0.28)
                    $pw.TextFrame2.TextRange.Font.Size = [single]28
                    $pw.ZOrder(0)
                }
                $m3d = Find-Shape $s "!!Hero3D"
                if ($null -ne $m3d) {
                    $m3d.Left = [single]($fx + $fw * 0.29); $m3d.Top = [single]($fy + $fh * 0.30)
                    $m3d.Width = [single]($fw * 0.42); $m3d.Height = [single]($fh * 0.66)
                    $m3d.ZOrder(0)
                } elseif ($null -ne (Find-Shape $s "!!poster_hero")) {
                    $ph = Find-Shape $s "!!poster_hero"
                    $ph.Left = [single]($fx + $fw * 0.28); $ph.Top = [single]($fy + $fh * 0.30)
                    $ph.Width = [single]($fw * 0.44); $ph.Height = [single]($fh * 0.66); $ph.ZOrder(0)
                }

                if ($sl._hollow_frame) {
                    $hfX = [single]($fx - ($fw * 0.0755))
                    $hfY = [single]($fy - ($fh * 0.152))
                    $hfW = [single]($fw / 0.86875)
                    $hfH = [single]($fh / 0.76667)
                    $hf = Add-Pic $s "!!hollow_frame" $sl._hollow_frame $hfX $hfY $hfW $hfH; $hf.ZOrder(0)
                }

                # Frameless gallery wall placard typography on serene dark wall
                Add-Text $s "!!pf_ovr" $sl.overline 62 148 380 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                $pfh = Add-Text $s "!!pf_h1" $sl.headline 62 178 396 102 $DisplayFont 34 $true $cWhite 1 1; Add-TextShadow $pfh
                Add-Text $s "!!pf_sub" $sl.subtitle 62 290 380 68 $BodyFont 13 $false $cMuted 1 1 | Out-Null
                Set-Morph $s ([single]2.2)
            }

            "real_portal_flythrough" {
                $stage = if ($sl.stage) { [string]$sl.stage } else { "exterior" }
                $keep = $carryBridge + @("!!portal_bg", "!!portal_inner", "!!portal_outer", "!!pt_ext_title", "!!pt_int_title", "!!pt_int_body", "!!HeroShadow", "!!Hero3D")
                $s = Start-MorphSlide $pres $prevSlide $keep

                if ($stage -eq "interior") {
                    $pbg = Add-Pic $s "!!portal_bg" "$AssetsDir\portal_bg.png" -48 -27 1056 594; $pbg.ZOrder(1)
                    $pout = Add-Pic $s "!!portal_outer" "$AssetsDir\portal_outer.png" -1900 -720 4620 2600
                    $pout.AlternativeText = "MORPH_ALPHA_ZERO"
                    $pout.ZOrder(0)
                    $pin = Add-Pic $s "!!portal_inner" "$AssetsDir\portal_inner.png" 0 0 960 540; $pin.ZOrder(0)

                    $extT = Find-Shape $s "!!pt_ext_title"
                    if ($null -ne $extT) { $extT.Top = [single]640 }

                    $ih1 = Add-Text $s "!!pt_int_title" $sl.headline 148 188 430 128 $DisplayFont 36 $true $cWhite 1 1; Add-TextShadow $ih1; $ih1.ZOrder(0)
                    $ibd = Add-Text $s "!!pt_int_body" $sl.body 148 320 410 82 $BodyFont 13 $false $cWhite 1 1; Add-TextShadow $ibd; $ibd.ZOrder(0)

                    if ($heroGlb -and ($sl.use_hero_3d -eq $true)) {
                        $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 596 468 284 48; $sh.ZOrder(0)
                        Ensure-Model3D $s "!!Hero3D" $heroGlb 600 188 276 312 10 -22 1 | Out-Null
                    }
                    Set-Morph $s ([single]2.1)
                } elseif ($stage -eq "horizon") {
                    $pbg = Add-Pic $s "!!portal_bg" "$AssetsDir\portal_bg.png" 0 0 960 540; $pbg.ZOrder(1)
                    $pin = Add-Pic $s "!!portal_inner" "$AssetsDir\portal_inner.png" -1820 -820 4400 2480
                    $pin.AlternativeText = "MORPH_ALPHA_ZERO"
                    $pin.ZOrder(0)
                    $ph1 = Add-Text $s "!!pt_int_title" $sl.headline 58 84 460 120 $DisplayFont 38 $true $cWhite 1 1; Add-TextShadow $ph1
                    $pbd = Add-Text $s "!!pt_int_body" $sl.body 58 214 430 96 $BodyFont 13 $false $cWhite 1 1; Add-TextShadow $pbd
                    Set-Morph $s ([single]2.0)
                } else {
                    $pbg = Add-Pic $s "!!portal_bg" "$AssetsDir\portal_bg.png" 0 0 960 540; $pbg.ZOrder(1)
                    $pin = Add-Pic $s "!!portal_inner" "$AssetsDir\portal_inner.png" 270 124 420 196; $pin.ZOrder(0)
                    $pout = Add-Pic $s "!!portal_outer" "$AssetsDir\portal_outer.png" 0 0 960 540; $pout.ZOrder(0)
                    $eh1 = Add-Text $s "!!pt_ext_title" $sl.headline 300 396 604 112 $DisplayFont 36 $true $cWhite 3 1; Add-TextShadow $eh1
                    if ($nextType -eq "real_portal_flythrough") {
                        Add-Text $s "!!pt_int_title" $nextSl.headline -560 188 430 128 $DisplayFont 36 $true $cWhite 1 1 | Out-Null
                        Add-Text $s "!!pt_int_body" $nextSl.body -680 320 410 82 $BodyFont 13 $false $cWhite 1 1 | Out-Null
                    }
                    Set-Morph $s ([single]2.0)
                }
            }

            "refracted_spotlight_dock" {
                $dockIdx = if ($null -ne $sl.dock_index) { [int]$sl.dock_index } else { 0 }
                $dockXs = @(66, 282, 498, 714); $gliderXs = @(54, 270, 486, 702)
                $keep = $carryBridge + @("!!GlassDock", "!!DockGlider", "!!DockThumb1", "!!DockThumb2", "!!DockThumb3", "!!DockThumb4", "!!sp_ring_0", "!!sp_ring_1", "!!sp_ring_2")
                if ($sl.use_hero_3d -eq $true) { $keep += @("!!HeroShadow", "!!Hero3D") }
                if ($prevType -eq "refracted_spotlight_dock") {
                    $prevIdx = $idx - 1
                    $keep += @("!!sp_bg_$idx", "!!sp_tag_$prevIdx", "!!sp_title_$prevIdx", "!!sp_sub_$prevIdx", "!!sp_body_$prevIdx", "!!DockLbl1", "!!DockLbl2", "!!DockLbl3", "!!DockLbl4", "!!DockSub1", "!!DockSub2", "!!DockSub3", "!!DockSub4")
                }
                $s = Start-MorphSlide $pres $prevSlide $keep
                $curBg = Add-Pic $s "!!sp_bg_$idx" "$AssetsDir\s$idx`_bg.png" 0 0 960 540; $curBg.ZOrder(1)

                if ($prevType -eq "refracted_spotlight_dock") {
                    $prevIdx = $idx - 1
                    foreach ($pair in @(@("!!sp_tag_$prevIdx", -280), @("!!sp_title_$prevIdx", -240), @("!!sp_sub_$prevIdx", -180), @("!!sp_body_$prevIdx", -140))) {
                        $osh = Find-Shape $s $pair[0]; if ($null -ne $osh) { $osh.Top = [single]$pair[1] }
                    }
                }
                if ($nextType -eq "refracted_spotlight_dock") {
                    $nextIdx = $idx + 1
                    $nbg = Add-Pic $s "!!sp_bg_$nextIdx" "$AssetsDir\s$nextIdx`_bg.png" 0 0 960 540
                    try { $nbg.PictureFormat.CropTop = [single](-360); $nbg.PictureFormat.CropBottom = [single]360; $nbg.Left = 0; $nbg.Top = 0; $nbg.Width = 960; $nbg.Height = 540 } catch {}
                    $nbg.ZOrder(1)
                    $nxSp = $nextSl.spotlight
                    Add-Text $s "!!sp_tag_$nextIdx" $nxSp.tag -460 48 520 20 $MonoFont 9.5 $true $cAcc3 1 3 | Out-Null
                    Add-Text $s "!!sp_title_$nextIdx" $nxSp.title -580 74 540 88 $DisplayFont 36 $true $cWhite 1 1 | Out-Null
                    Add-Text $s "!!sp_sub_$nextIdx" $nxSp.subtitle -700 168 520 26 $SemiFont 13 $false $cWhite 1 3 | Out-Null
                    Add-Text $s "!!sp_body_$nextIdx" $nxSp.body -820 202 490 104 $BodyFont 13 $false $cWhite 1 1 | Out-Null
                }

                $sp = $sl.spotlight
                Add-Text $s "!!sp_tag_$idx" $sp.tag 56 48 520 20 $MonoFont 9.5 $true $cAcc3 1 3 | Out-Null
                $spt = Add-Text $s "!!sp_title_$idx" $sp.title 56 74 540 88 $DisplayFont 36 $true $cWhite 1 1; Add-TextShadow $spt
                Add-Text $s "!!sp_sub_$idx" $sp.subtitle 56 168 520 26 $SemiFont 13 $false $cWhite 1 3 | Out-Null
                $spb = Add-Text $s "!!sp_body_$idx" $sp.body 56 202 490 104 $BodyFont 13 $false $cWhite 1 1; Add-TextShadow $spb

                if ($heroGlb -and ($sl.use_hero_3d -eq $true)) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 644 304 268 46; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 648 48 248 272 10 -20 2 | Out-Null
                }

                $refractDock = "$AssetsDir\s$idx`_dock_refract.png"
                $dockPath = if (Test-Path $refractDock) { $refractDock } else { "$AssetsDir\dock_bg.png" }
                $gd = Find-Shape $s "!!GlassDock"
                if ($null -ne $gd -and $prevType -eq "refracted_spotlight_dock") { $gd.Delete() }
                $gdk = Add-Pic $s "!!GlassDock" $dockPath 40 386 880 128; $gdk.ZOrder(0)
                $gld = Add-Pic $s "!!DockGlider" "$AssetsDir\dock_glider.png" $gliderXs[[Math]::Min(3, [Math]::Max(0, $dockIdx))] 396 208 108; $gld.ZOrder(0)

                $dockItems = @($sl.dock)
                for ($di = 0; $di -lt 4; $di++) {
                    $it = if ($di -lt $dockItems.Count) { $dockItems[$di] } else { @{ title = "Item $($di+1)"; subtitle = "" } }
                    $dn = $di + 1; $dx = $dockXs[$di]
                    $th = Add-Pic $s "!!DockThumb$dn" "$AssetsDir\thumb_$dn.png" $dx 416 76 60; if ($th) { $th.ZOrder(0) }
                    $tl = Add-Text $s "!!DockLbl$dn" $it.title ($dx + 84) 418 112 24 $DisplayFont 13 $true $cWhite 1 3; $tl.ZOrder(0)
                    $ts = Add-Text $s "!!DockSub$dn" $it.subtitle ($dx + 84) 444 112 24 $BodyFont 10.5 $false $cWhite 1 1; $ts.ZOrder(0)
                }
                for ($pi = 0; $pi -lt 3; $pi++) {
                    $isAct = ($pi -eq ($dockIdx % 3))
                    $sz = if ($isAct) { [single]10.5 } else { [single]7.0 }
                    $px = [single](928.0 - ($sz / 2.0))
                    $py = [single](210 + $pi * 20)
                    $ring = Find-Shape $s "!!sp_ring_$pi"
                    if ($null -eq $ring) { $ring = $s.Shapes.AddShape(9, $px, $py, $sz, $sz); $ring.Name = "!!sp_ring_$pi" }
                    else { $ring.Left = $px; $ring.Top = $py; $ring.Width = $sz; $ring.Height = $sz }
                    if ($isAct) { $ring.Fill.Visible = -1; $ring.Fill.Solid(); $ring.Fill.ForeColor.RGB = $cWhite } else { $ring.Fill.Visible = 0 }
                    $ring.Line.Visible = -1; $ring.Line.ForeColor.RGB = $cWhite; $ring.Line.Weight = [single]1.2
                    $ring.ZOrder(0)
                }
                Set-Morph $s ([single]2.0)
            }

            "parchment_3d_split" {
                $keep = $carryBridge + @("!!SplitLower", "!!GiantWord", "!!HeroShadow", "!!Hero3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bg = Add-Pic $s "!!StageBg" "$AssetsDir\bg_base.png" 0 0 960 540; $bg.ZOrder(1)
                $lowPath = "$AssetsDir\s$idx`_split_lower.png"
                if (Test-Path $lowPath) { $lw = Add-Pic $s "!!SplitLower" $lowPath 0 272 960 268; $lw.ZOrder(0) }

                $gwTxt = if ($sl.giant_word) { [string]$sl.giant_word } else { "STUDIO" }
                $gw = Add-Text $s "!!GiantWord" $gwTxt 30 18 900 168 $DisplayFont 118 $true $cAcc1 2 3; $gw.ZOrder(0)
                Add-Text $s "!!s2_pl" $sl.left_label 42 186 248 18 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "!!s2_pt" $sl.left_text 42 206 248 62 $SemiFont 11.5 $false $cInk 1 1 | Out-Null
                Add-Text $s "!!s2_sl" $sl.right_label 670 186 248 18 $MonoFont 9.5 $true $cAcc1 3 3 | Out-Null
                Add-Text $s "!!s2_st" $sl.right_text 670 206 248 62 $SemiFont 11.5 $false $cInk 3 1 | Out-Null

                $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 334 468 292 54; $sh.ZOrder(0)
                if ($heroGlb) {
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 318 96 324 396 8 -16 0 | Out-Null
                }
                Set-Morph $s ([single]2.1)
            }

            "vertical_world_elevator" {
                $stop = if ($sl.elevator_stop) { [string]$sl.elevator_stop } else { "upper_sky" }
                $gwTxt = if ($sl.giant_word) { [string]$sl.giant_word } else { "HORIZON" }
                $keep = $carryBridge + @("!!elev_bg", "!!elev_word", "!!elev_fg", "!!ev_up_h1", "!!ev_rcard_0", "!!ev_rcard_1", "!!ev_rcard_2", "!!HeroShadow", "!!Hero3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $ebgP = if (Test-Path "$AssetsDir\elev_bg.png") { "$AssetsDir\elev_bg.png" } else { "$AssetsDir\bg_base.png" }

                if ($stop -eq "lower_ground") {
                    $ebg = Add-Pic $s "!!elev_bg" $ebgP 0 -520 960 1080; $ebg.ZOrder(1)
                    $ew = Add-Text $s "!!elev_word" $gwTxt 40 286 880 170 $HeadingFont 128 $true $cWhite 2 3; Add-TextShadow $ew; $ew.ZOrder(0)
                    if (Test-Path "$AssetsDir\elev_fg.png") { $efg = Add-Pic $s "!!elev_fg" "$AssetsDir\elev_fg.png" 0 -520 960 1080; $efg.ZOrder(0) }
                    $upH = Find-Shape $s "!!ev_up_h1"; if ($null -ne $upH) { $upH.Top = [single]-260 }

                    if ($heroGlb) {
                        $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 54 456 336 56; $sh.ZOrder(0)
                        Ensure-Model3D $s "!!Hero3D" $heroGlb 58 112 328 368 14 26 -2 | Out-Null
                    }
                    $eh2 = Add-Text $s "!!ev_h2" $sl.headline 444 52 468 64 $DisplayFont 26 $true $cWhite 1 1; Add-TextShadow $eh2
                    $cards = @($sl.cards)
                    for ($ci = 0; $ci -lt [Math]::Min(3, $cards.Count); $ci++) {
                        $cy = [single](142 + $ci * 118); $crd = $cards[$ci]
                        $rc = Add-Pic $s "!!ev_rcard_$ci" "$AssetsDir\card_spec_$($ci + 1).png" 444 $cy 468 102; $rc.ZOrder(0)
                        Add-Text $s "ev_rmet_$ci" $crd.value 464 ($cy + 16) 124 66 $DisplayFont 28 $true $cAcc1 2 3 | Out-Null
                        Add-Text $s "ev_rttl_$ci" $crd.title 596 ($cy + 16) 294 26 $DisplayFont 14 $true $cInk 1 3 | Out-Null
                        Add-Text $s "ev_rdsc_$ci" $crd.desc 596 ($cy + 44) 294 44 $BodyFont 11 $false $cMuted 1 1 | Out-Null
                    }
                    Set-Morph $s ([single]2.2)
                } else {
                    # Delete any mid-air Hero3D on upper_sky unless next slide needs it off-canvas
                    $oldH = Find-Shape $s "!!Hero3D"; if ($null -ne $oldH -and $nextType -ne "vertical_world_elevator") { $oldH.Delete() }
                    $oldS = Find-Shape $s "!!HeroShadow"; if ($null -ne $oldS) { $oldS.Delete() }

                    $ebg = Add-Pic $s "!!elev_bg" $ebgP 0 0 960 1080; $ebg.ZOrder(1)
                    $ew = Add-Text $s "!!elev_word" $gwTxt 40 68 880 168 $HeadingFont 132 $true $cWhite 2 3; Add-TextShadow $ew; $ew.ZOrder(0)
                    if (Test-Path "$AssetsDir\elev_fg.png") { $efg = Add-Pic $s "!!elev_fg" "$AssetsDir\elev_fg.png" 0 0 960 1080; $efg.ZOrder(0) }

                    # Frameless white editorial copy directly over dark mountain slope (zero card_panel!)
                    Add-Text $s "!!ev_up_ovr" $sl.overline 56 334 420 20 $MonoFont 9.5 $true $cWhite 1 3 | Out-Null
                    $euh = Add-Text $s "!!ev_up_h1" $sl.headline 56 358 460 72 $DisplayFont 30 $true $cWhite 1 1; Add-TextShadow $euh
                    $eub = Add-Text $s "!!ev_up_body" $sl.body 56 434 440 68 $BodyFont 12.5 $false $cWhite 1 1; Add-TextShadow $eub

                    if ($nextType -eq "vertical_world_elevator") {
                        for ($ci = 0; $ci -lt 3; $ci++) {
                            Add-Pic $s "!!ev_rcard_$ci" "$AssetsDir\card_spec_$($ci + 1).png" (1060 + $ci * 160) (142 + $ci * 118) 468 102 | Out-Null
                        }
                        if ($heroGlb) {
                            Ensure-Model3D $s "!!Hero3D" $heroGlb -380 140 328 368 10 -18 2 | Out-Null
                        }
                    }
                    Set-Morph $s ([single]2.1)
                }
            }

            "rotational_drift_3d" {
                $keep = $carryBridge + @("!!HeroShadow", "!!Hero3D", "!!md_card_0", "!!md_card_1", "!!md_card_2")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bgP = if ($sl._bg_image) { [string]$sl._bg_image } else { "$AssetsDir\bg_base.png" }
                $bg = Add-Pic $s "!!StageBg" $bgP 0 0 960 540; $bg.ZOrder(1)

                Add-Text $s "!!md_ovr" $sl.overline 48 26 520 18 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "!!md_h1" $sl.headline 48 48 470 74 $DisplayFont 26 $true $cInk 1 1 | Out-Null
                if ($heroGlb) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 496 434 404 62; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 498 58 392 392 18 -34 -4 | Out-Null
                }
                $callouts = @($sl.callouts); $cCols = @($cAcc1, $cAcc2, $cAcc3)
                for ($ci = 0; $ci -lt [Math]::Min(3, $callouts.Count); $ci++) {
                    $cy = [single](144 + $ci * 116); $ac = $cCols[$ci]
                    $mc = Add-Pic $s "!!md_card_$ci" "$AssetsDir\card_spec_$($ci + 1).png" 48 $cy 384 96; $mc.ZOrder(0)
                    Add-Text $s "md_ttl_$ci" $callouts[$ci].title 68 ($cy + 14) 344 22 $DisplayFont 13.5 $true $ac 1 3 | Out-Null
                    Add-Text $s "md_dsc_$ci" $callouts[$ci].desc 68 ($cy + 40) 344 44 $BodyFont 11 $false $cInk 1 1 | Out-Null
                    Add-Rect $s "!!md_line_$ci" 432 ($cy + 47) 54 ([single]1.5) $ac 0 -1 0 0 | Out-Null
                }
                Set-Morph $s ([single]2.0)
            }

            "swiss_bento_morph" {
                $keep = $carryBridge + @("!!HeroShadow", "!!Hero3D", "!!kb_hero", "!!kb_tbg_0", "!!kb_tbg_1", "!!kb_tbg_2")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bg = Add-Pic $s "!!StageBg" "$AssetsDir\bg_base.png" 0 0 960 540; $bg.ZOrder(1)
                Add-Text $s "!!kb_ovr" $sl.overline 48 24 620 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "!!kb_h1" $sl.headline 48 46 660 54 $DisplayFont 25 $true $cInk 1 1 | Out-Null

                $hero = $sl.hero_tile
                $kh = Add-Pic $s "!!kb_hero" "$AssetsDir\card_bento_strip.png" 46 106 588 154; $kh.ZOrder(0)
                Add-Text $s "kb_htag" $hero.label 74 120 180 18 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "kb_hval" $hero.value 74 144 180 84 $DisplayFont 44 $true $cInk 1 1 | Out-Null
                Add-Rect $s "kb_vdiv" 262 128 1 108 $cBorder 0 -1 0 0 | Out-Null
                Add-Text $s "kb_hsub" $hero.title 280 124 334 30 $DisplayFont 16.5 $true $cInk 1 3 | Out-Null
                Add-Text $s "kb_hdsc" $hero.desc 280 158 334 62 $BodyFont 11.5 $false $cMuted 1 1 | Out-Null

                if ($heroGlb) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 654 236 256 42; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 668 42 232 214 14 -28 3 | Out-Null
                }

                $tiles = @($sl.tiles); $bCols = @($cAcc1, $cAcc2, $cAcc3)
                for ($ti = 0; $ti -lt [Math]::Min(3, $tiles.Count); $ti++) {
                    $t = $tiles[$ti]; $tx = [single](46 + $ti * 296)
                    $kt = Add-Pic $s "!!kb_tbg_$ti" "$AssetsDir\card_bento_tile_$($ti + 1).png" $tx 276 276 222; $kt.ZOrder(0)
                    Add-Text $s "kb_tlbl_$ti" $t.label ($tx + 22) 298 232 20 $MonoFont 9.5 $true $bCols[$ti % 3] 1 3 | Out-Null
                    Add-Text $s "kb_tval_$ti" $t.value ($tx + 22) 324 232 62 $DisplayFont 38 $true $cInk 1 3 | Out-Null
                    Add-Text $s "kb_tdsc_$ti" $t.desc ($tx + 22) 396 232 82 $BodyFont 11.5 $false $cMuted 1 1 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "timeline_rail_morph" {
                $stages = @($sl.stages); $nSt = [Math]::Max(1, [Math]::Min(4, $stages.Count))
                $actIdx = [Math]::Max(0, [Math]::Min($nSt - 1, [int]$sl.active_index))
                $keep = $carryBridge + @("!!HeroShadow", "!!Hero3D", "!!tl_rail", "!!tl_drawer")
                if ($prevType -eq "timeline_rail_morph") { $keep += "!!tl_progress"; for ($ti = 0; $ti -lt $nSt; $ti++) { $keep += "!!tl_node_$ti" } }
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bg = Add-Pic $s "!!StageBg" "$AssetsDir\bg_base.png" 0 0 960 540; $bg.ZOrder(1)

                Add-Text $s "!!tl_ovr" $sl.overline 52 24 620 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "!!tl_h1" $sl.headline 52 46 760 54 $DisplayFont 25 $true $cInk 1 1 | Out-Null
                Add-Rect $s "!!tl_rail" 76 144 808 3 $cBorder 0 -1 0 0 | Out-Null
                $centers = @(); for ($ti = 0; $ti -lt $nSt; $ti++) { $centers += if ($nSt -eq 1) { [single]480 } else { [single](146 + 668 * ($ti / ($nSt - 1))) } }
                Add-Rect $s "!!tl_progress" 76 144 ([Math]::Max([single]24, $centers[$actIdx] - 76)) 3 $cAcc1 0 -1 0 0 | Out-Null
                for ($ti = 0; $ti -lt $nSt; $ti++) {
                    $stg = $stages[$ti]; $cx = $centers[$ti]
                    $nw = if ($ti -eq $actIdx) { [single]172 } else { [single]148 }
                    $bgF = if ($ti -eq $actIdx) { $cAcc1 } else { $cSurface }
                    $fgT = if ($ti -eq $actIdx) { $cWhite } else { $cMuted }
                    Add-CapsulePill $s "!!tl_node_$ti" "$($stg.code)  $($stg.label)" ([single]($cx - $nw / 2)) ([single]128) $nw ([single]34) $bgF 0 $fgT 9.5 $MonoFont | Out-Null
                }
                if ($heroGlb) {
                    $rotYs = @(22, -24, 38, -16)
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 34 448 272 50; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 46 192 250 274 12 ([single]$rotYs[$actIdx % 4]) 2 | Out-Null
                }
                $cur = $stages[$actIdx]
                $drw = Add-Pic $s "!!tl_drawer" "$AssetsDir\card_panel.png" 316 204 596 292; $drw.ZOrder(0)
                Add-Text $s "tl_metric" $cur.metric 344 240 176 76 $DisplayFont 42 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "tl_mlbl" $cur.metric_label 344 322 176 36 $SemiFont 11 $false $cInk 1 1 | Out-Null
                Add-Text $s "tl_sh1" $cur.title 544 226 346 52 $DisplayFont 18.5 $true $cInk 1 1 | Out-Null
                Add-Text $s "tl_sum" $cur.summary 544 284 346 110 $BodyFont 12 $false $cMuted 1 1 | Out-Null
                Set-Morph $s ([single]1.8)
            }

            "chart_infographic" {
                $keep = $carryBridge + @("!!HeroShadow", "!!Hero3D", "!!ch_img", "!!ch_hlbg")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bg = Add-Pic $s "!!StageBg" "$AssetsDir\bg_base.png" 0 0 960 540; $bg.ZOrder(1)
                Add-Text $s "!!ch_ovr" $sl.overline 46 24 600 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "!!ch_h1" $sl.headline 46 46 600 54 $DisplayFont 25 $true $cInk 1 1 | Out-Null
                $chi = Add-Pic $s "!!ch_img" "$AssetsDir\s$idx`_chart.png" 46 110 584 384; $chi.ZOrder(0)
                if ($heroGlb) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 656 244 248 44; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 672 38 220 222 16 28 -2 | Out-Null
                }
                $hl = $sl.highlight
                $hlb = Add-Pic $s "!!ch_hlbg" "$AssetsDir\card_spec_1.png" 648 278 268 104; $hlb.ZOrder(0)
                Add-Text $s "ch_hlv" "$($hl.value)  $($hl.label)" 670 290 232 30 $DisplayFont 18 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "ch_hld" $hl.desc 670 324 232 48 $BodyFont 10.5 $false $cInk 1 1 | Out-Null
                Set-Morph $s ([single]1.8)
            }

            "architecture_blueprint" {
                $keep = $carryBridge + @("!!HeroShadow", "!!Hero3D", "!!ar_img")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bg = Add-Pic $s "!!StageBg" "$AssetsDir\bg_base.png" 0 0 960 540; $bg.ZOrder(1)
                Add-Text $s "!!ar_ovr" $sl.overline 42 22 620 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "!!ar_h1" $sl.headline 42 46 630 52 $DisplayFont 24 $true $cInk 1 1 | Out-Null
                $ari = Add-Pic $s "!!ar_img" "$AssetsDir\s$idx`_arch.png" 42 104 624 398; $ari.ZOrder(0)
                if ($heroGlb) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 686 244 236 42; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 698 36 212 222 22 -34 -3 | Out-Null
                }
                $pillars = @($sl.pillars)
                for ($pi = 0; $pi -lt [Math]::Min(2, $pillars.Count); $pi++) {
                    $pl = $pillars[$pi]; $sy = [single](276 + $pi * 116)
                    Add-Pic $s "!!ar_cbg_$pi" "$AssetsDir\card_spec_$($pi + 1).png" 682 $sy 236 106 | Out-Null
                    Add-Text $s "ar_ctag_$pi" "$($pl.tag)  //  $($pl.title)" 702 ($sy + 12) 204 22 $DisplayFont 12 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "ar_cdsc_$pi" $pl.desc 702 ($sy + 38) 204 58 $BodyFont 10.5 $false $cInk 1 1 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "closing_finale" {
                $keep = $carryBridge + @("!!HeroShadow", "!!Hero3D", "!!fs_h1", "!!fs_rule")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bgP = if ($sl._bg_image) { [string]$sl._bg_image } else { "$AssetsDir\bg_base.png" }
                $bg = Add-Pic $s "!!StageBg" $bgP 0 0 960 540; $bg.ZOrder(1)
                $txtCol = if ($sl._bg_image) { $cWhite } else { $cInk }
                $mutCol = if ($sl._bg_image) { $cWhite } else { $cMuted }

                Add-Text $s "!!fs_ovr" $sl.overline 54 58 460 22 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "!!fs_h1" $sl.headline 54 92 496 118 $DisplayFont 36 $true $txtCol 1 1 | Out-Null
                Add-Rect $s "!!fs_rule" 54 222 96 ([single]2.5) $cAcc1 0 -1 0 0 | Out-Null
                Add-Text $s "!!fs_sub" $sl.summary 54 242 480 76 $BodyFont 13.5 $false $mutCol 1 1 | Out-Null
                $metrics = @($sl.metrics)
                for ($mi = 0; $mi -lt [Math]::Min(3, $metrics.Count); $mi++) {
                    $mx = [single](54 + $mi * 164)
                    Add-Rect $s "!!fs_mline_$mi" $mx 344 142 2 $cAcc1 0 -1 0 0 | Out-Null
                    Add-Text $s "!!fs_mv_$mi" $metrics[$mi].value $mx 354 142 42 $DisplayFont 28 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "!!fs_ml_$mi" $metrics[$mi].label $mx 398 142 22 $MonoFont 9 $false $mutCol 1 1 | Out-Null
                }
                if ($heroGlb) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 536 430 352 56; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 548 72 336 368 10 -20 1 | Out-Null
                }
                Set-Morph $s ([single]2.0)
            }

            default { throw "Unsupported slide type '$st' at index $idx" }
        }

        # --- Post-stage outgoing portal/elevator layers with MORPH_ALPHA_ZERO optical dissolve ---
        if ($prevType -eq "real_portal_flythrough" -and $st -ne "real_portal_flythrough") {
            $pbOut = Find-Shape $s "!!portal_bg"
            if ($null -ne $pbOut) { $pbOut.Left = -96; $pbOut.Top = -54; $pbOut.Width = 1152; $pbOut.Height = 648; $pbOut.AlternativeText = "MORPH_ALPHA_ZERO"; $pbOut.ZOrder(0) }
            $piOut = Find-Shape $s "!!portal_inner"
            if ($null -ne $piOut) { $piOut.Left = -1820; $piOut.Top = -820; $piOut.Width = 4400; $piOut.Height = 2480; $piOut.AlternativeText = "MORPH_ALPHA_ZERO"; $piOut.ZOrder(0) }
        }
        if ($prevType -eq "vertical_world_elevator" -and $st -ne "vertical_world_elevator") {
            $ebOut = Find-Shape $s "!!elev_bg"
            if ($null -ne $ebOut) { $ebOut.Left = 0; $ebOut.Top = -980; $ebOut.Width = 960; $ebOut.Height = 1080; $ebOut.AlternativeText = "MORPH_ALPHA_ZERO" }
            $efOut = Find-Shape $s "!!elev_fg"
            if ($null -ne $efOut) { $efOut.Left = 0; $efOut.Top = -980; $efOut.Width = 960; $efOut.Height = 1080; $efOut.AlternativeText = "MORPH_ALPHA_ZERO" }
        }

        # --- Universal Pre-Staging Bridge for the NEXT slide ---
        if ($null -ne $nextSl) {
            if ($nextType -eq "editorial_print_overlay" -and $st -ne "editorial_print_overlay") {
                Add-Pic $s "!!ep_drips" "$AssetsDir\ep_drips.png" 0 -180 960 120 | Out-Null
                $portNext = if ($nextSl._portrait_cutout) { [string]$nextSl._portrait_cutout } elseif (Test-Path "$AssetsDir\hero_cutout_auto.png") { "$AssetsDir\hero_cutout_auto.png" } else { "" }
                if ($portNext) { Add-Pic $s "!!ep_port" $portNext -390 86 344 444 | Out-Null }
                Add-Pic $s "!!ep_cloud" "$AssetsDir\ep_paper_cloud.png" 1020 118 528 356 | Out-Null
            } elseif ($nextType -eq "odometer_counter") {
                $nextBg = if ($nextSl.bg_hex) { HexToBgr $nextSl.bg_hex } else { $cObsidian }
                Add-Rect $s "!!odo_curtain" -980 0 960 540 $nextBg 0 -1 0 0 | Out-Null
                $nextChars = ([string]$nextSl.target_value).ToCharArray()
                $totW = 0; foreach ($ch in $nextChars) { $totW += if ($ch -eq ' ') { 24 } elseif ($ch -match '\d') { 50 } else { 32 } }
                $cx = [single](-980 + (480 - ($totW / 2)))
                for ($ci = 0; $ci -lt $nextChars.Count; $ci++) {
                    $ch = $nextChars[$ci]
                    if ($ch -eq ' ') { $cx += 24; continue }
                    if ($ch -match '\d') {
                        Add-OdometerStrip $s "!!odo_col_$ci" $cx ([single](218 + (($ci % 3) - 9) * 86)) 50 $cWhite $HeadingFont | Out-Null
                        $cx += 50
                    } else { $cx += 32 }
                }
                Add-Rect $s "!!odo_mask_top" -984 -4 968 222 $nextBg 0 -1 0 0 | Out-Null
                Add-Rect $s "!!odo_mask_bot" -984 304 968 242 $nextBg 0 -1 0 0 | Out-Null
                Add-Pic $s "!!odo_sil" "$AssetsDir\odo_silhouette.png" -980 378 960 162 | Out-Null
            } elseif ($nextType -eq "flyby_wipe") {
                Add-Pic $s "!!flyby_cloud" "$AssetsDir\flyby_cloud.png" -860 320 760 440 | Out-Null
                if ($nextSl._flyby_cutout) {
                    $foPre = Add-Pic $s "!!flyby_obj" $nextSl._flyby_cutout -480 400 440 280
                    $foPre.Rotation = [single]-8
                }
                if ($heroGlb -and ($nextSl.use_hero_3d -ne $false) -and $null -eq (Find-Shape $s "!!Hero3D")) {
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 1040 92 336 360 8 18 -4 | Out-Null
                }
            } elseif ($nextType -eq "cutout_map_stage") {
                if ($nextSl._map_image) { Add-Pic $s "!!cm_map" $nextSl._map_image -560 54 520 432 | Out-Null }
                Add-Rect $s "!!cm_div" 1020 68 ([single]1.5) 404 $cBorder 0 -1 0 0 | Out-Null
                Add-Text $s "!!cm_h1" $nextSl.headline 1040 100 294 84 $DisplayFont 30 $true $cAcc1 1 1 | Out-Null
            } elseif ($nextType -eq "poster_depth_reveal") {
                Add-Rect $s "!!poster_band" 348 -560 264 540 $cAcc1 0 -1 0 0 | Out-Null
                $ngw = if ($nextSl.giant_word) { [string]$nextSl.giant_word } else { "2026" }
                Add-Text $s "!!poster_word" $ngw 60 -180 840 138 $HeadingFont 116 $true $cWhite 2 3 | Out-Null
                if ($nextSl._hero_cutout) { Add-Pic $s "!!poster_hero" $nextSl._hero_cutout 284 590 392 368 | Out-Null }
                elseif ($heroGlb -and $null -eq (Find-Shape $s "!!Hero3D")) { Ensure-Model3D $s "!!Hero3D" $heroGlb 296 580 368 364 16 26 -2 | Out-Null }
            } elseif ($nextType -eq "parchment_3d_split" -and $st -ne "parchment_3d_split") {
                $lowNext = "$AssetsDir\s$($idx + 1)_split_lower.png"
                if (Test-Path $lowNext) { Add-Pic $s "!!SplitLower" $lowNext 0 580 960 268 | Out-Null }
                Add-Text $s "!!GiantWord" $nextSl.giant_word 30 -210 900 168 $DisplayFont 118 $true $cAcc1 2 3 | Out-Null
                if ($heroGlb -and $null -eq (Find-Shape $s "!!Hero3D")) { Ensure-Model3D $s "!!Hero3D" $heroGlb 318 580 324 396 14 -28 2 | Out-Null }
            } elseif ($nextType -eq "vertical_world_elevator" -and $st -ne "vertical_world_elevator") {
                if (Test-Path "$AssetsDir\elev_bg.png") { Add-Pic $s "!!elev_bg" "$AssetsDir\elev_bg.png" 0 580 960 1080 | Out-Null }
                Add-Text $s "!!elev_word" $nextSl.giant_word 40 760 880 168 $HeadingFont 132 $true $cWhite 2 3 | Out-Null
                if (Test-Path "$AssetsDir\elev_fg.png") { Add-Pic $s "!!elev_fg" "$AssetsDir\elev_fg.png" 0 580 960 1080 | Out-Null }
            } elseif ($nextType -eq "refracted_spotlight_dock" -and $st -ne "refracted_spotlight_dock") {
                $dockNext = "$AssetsDir\s$($idx + 1)_dock_refract.png"
                $dpN = if (Test-Path $dockNext) { $dockNext } else { "$AssetsDir\dock_bg.png" }
                Add-Pic $s "!!GlassDock" $dpN 40 620 880 128 | Out-Null
                Add-Pic $s "!!DockGlider" "$AssetsDir\dock_glider.png" 54 680 208 108 | Out-Null
                for ($dn = 1; $dn -le 4; $dn++) {
                    if (Test-Path "$AssetsDir\thumb_$dn.png") {
                        Add-Pic $s "!!DockThumb$dn" "$AssetsDir\thumb_$dn.png" (66 + ($dn - 1) * 216) (650 + $dn * 55) 76 60 | Out-Null
                    }
                }
            } elseif ($nextType -eq "rotational_drift_3d" -and $st -ne "rotational_drift_3d") {
                for ($mi = 0; $mi -lt 3; $mi++) {
                    Add-Pic $s "!!md_card_$mi" "$AssetsDir\card_spec_$($mi + 1).png" (-460 - $mi * 140) (144 + $mi * 116) 384 96 | Out-Null
                }
            } elseif ($nextType -eq "swiss_bento_morph" -and $st -ne "swiss_bento_morph") {
                Add-Pic $s "!!kb_hero" "$AssetsDir\card_bento_strip.png" -640 106 588 154 | Out-Null
                for ($ti = 0; $ti -lt 3; $ti++) {
                    Add-Pic $s "!!kb_tbg_$ti" "$AssetsDir\card_bento_tile_$($ti + 1).png" (46 + $ti * 296) (580 + $ti * 70) 276 222 | Out-Null
                }
            } elseif ($nextType -eq "chart_infographic" -and $st -ne "chart_infographic") {
                Add-Pic $s "!!ch_img" "$AssetsDir\s$($idx + 1)_chart.png" -640 110 584 384 | Out-Null
                Add-Pic $s "!!ch_hlbg" "$AssetsDir\card_spec_1.png" 1020 278 268 104 | Out-Null
            } elseif ($nextType -eq "timeline_rail_morph" -and $st -ne "timeline_rail_morph") {
                Add-Rect $s "!!tl_rail" -860 144 808 3 $cBorder 0 -1 0 0 | Out-Null
                Add-Pic $s "!!tl_drawer" "$AssetsDir\card_panel.png" 316 590 596 292 | Out-Null
            } elseif ($nextType -eq "architecture_blueprint" -and $st -ne "architecture_blueprint") {
                Add-Pic $s "!!ar_img" "$AssetsDir\s$($idx + 1)_arch.png" -680 104 624 398 | Out-Null
            } elseif ($nextType -eq "closing_finale" -and $st -ne "closing_finale") {
                Add-Text $s "!!fs_h1" $nextSl.headline -560 92 496 118 $DisplayFont 36 $true $cInk 1 1 | Out-Null
                Add-Rect $s "!!fs_rule" -560 222 96 ([single]2.5) $cAcc1 0 -1 0 0 | Out-Null
                if ($heroGlb -and $null -eq (Find-Shape $s "!!Hero3D")) { Ensure-Model3D $s "!!Hero3D" $heroGlb 1040 72 336 368 14 24 -2 | Out-Null }
            }
        }

        $prevSlide = $s
        $prevType  = $st
        $prevSl    = $sl
    }

    if (Test-Path $OutputPptx) { Remove-Item $OutputPptx -Force -ErrorAction SilentlyContinue }
    $pres.SaveAs($OutputPptx)
    Write-Host ("Saved True-Morph v4.0.0 deck -> " + $OutputPptx)

    if (-not $SkipExport) {
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $pngPath = Join-Path $ExportDir "slide_$i.png"
            $pres.Slides.Item($i).Export($pngPath, "PNG", 1920, 1080)
            Write-Host "Exported slide $i -> $pngPath"
        }
        if (-not [string]::IsNullOrWhiteSpace($CopyTo)) {
            Copy-Item -Path $OutputPptx -Destination $CopyTo -Force
            Write-Host ("Copied deck -> " + $CopyTo)
        }
    }
} finally {
    if ($null -ne $pres) { try { $pres.Close() } catch {} }
    try { $ppt.Quit() } catch {}
    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
