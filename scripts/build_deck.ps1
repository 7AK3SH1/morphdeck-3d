# Copyright 2026 Mohamed S. Sissoko (7AK3SH1)
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).

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
        $pres = $ppt.Presentations.Open($OutputPptx, -1, 0, 0) # ReadOnly, msoFalse
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $pngPath = Join-Path $ExportDir "slide_$i.png"
            $pres.Slides.Item($i).Export($pngPath, "PNG", 1920, 1080)
            Write-Host "Exported lit slide $i -> $pngPath"
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

if ([string]::IsNullOrWhiteSpace($SpecJson) -or [string]::IsNullOrWhiteSpace($AssetsDir)) {
    throw "SpecJson and AssetsDir are required unless -ExportOnly is specified."
}

$resolvedPath = Join-Path $AssetsDir "theme_resolved.json"
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
    $slide.SlideShowTransition.Speed = 1 # ppTransitionSpeedSlow
}

function Add-Bg($slide, [string]$imgPath) {
    $existing = Find-Shape $slide "bg_slide"
    if ($null -ne $existing) { $existing.Delete() }
    $bg = $slide.Shapes.AddPicture($imgPath, 0, 1, [single]0, [single]0, [single]960, [single]540)
    $bg.Name = "bg_slide"
    $bg.ZOrder(1) # msoSendToBack
    return $bg
}

function Add-Pic($slide, [string]$name, [string]$imgPath, [single]$x, [single]$y, [single]$w, [single]$h) {
    $existing = Find-Shape $slide $name
    if ($null -ne $existing) {
        try {
            $existing.PictureFormat.CropTop = [single]0
            $existing.PictureFormat.CropBottom = [single]0
        } catch {}
        $existing.Left = $x; $existing.Top = $y; $existing.Width = $w; $existing.Height = $h
        return $existing
    }
    $sh = $slide.Shapes.AddPicture($imgPath, 0, 1, $x, $y, $w, $h)
    $sh.Name = $name
    return $sh
}

function Add-TextShadow($tb) {
    try {
        $sh = $tb.TextFrame2.TextRange.Font.Shadow
        $sh.Visible = -1
        $sh.Blur = [single]6
        $sh.OffsetX = [single]0
        $sh.OffsetY = [single]3
        $sh.Transparency = [single]0.48
    } catch {}
}

function Add-Rect($slide, [string]$name, [single]$x, [single]$y, [single]$w, [single]$h, [int]$fillBgr, [single]$fillTrans = 0, [int]$borderBgr = -1, [single]$borderWeight = 0, [single]$cornerAdj = 0.0) {
    $existing = Find-Shape $slide $name
    if ($null -ne $existing) {
        $existing.Left = $x; $existing.Top = $y; $existing.Width = $w; $existing.Height = $h
        $existing.Fill.Solid()
        $existing.Fill.ForeColor.RGB = $fillBgr
        $existing.Fill.Transparency = $fillTrans
        if ($borderBgr -ge 0 -and $borderWeight -gt 0) {
            $existing.Line.Visible = -1
            $existing.Line.ForeColor.RGB = $borderBgr
            $existing.Line.Weight = $borderWeight
        } else {
            $existing.Line.Visible = 0
        }
        return $existing
    }
    $shapeType = if ($cornerAdj -gt 0.0) { 5 } else { 1 }
    $sh = $slide.Shapes.AddShape($shapeType, $x, $y, $w, $h)
    $sh.Name = $name
    if ($cornerAdj -gt 0.0) { $sh.Adjustments.Item(1) = $cornerAdj }
    $sh.Fill.Solid()
    $sh.Fill.ForeColor.RGB = $fillBgr
    $sh.Fill.Transparency = $fillTrans
    if ($borderBgr -ge 0 -and $borderWeight -gt 0) {
        $sh.Line.Visible = -1
        $sh.Line.ForeColor.RGB = $borderBgr
        $sh.Line.Weight = $borderWeight
    } else {
        $sh.Line.Visible = 0
    }
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
    $tf.WordWrap = -1
    $tf.AutoSize = 0
    $tf.MarginLeft = [single]2
    $tf.MarginRight = [single]2
    $tf.MarginTop = [single]2
    $tf.MarginBottom = [single]2
    $tf.VerticalAnchor = $vAnchor
    $tr = $tf.TextRange
    $tr.Text = if ($null -eq $text) { "" } else { [string]$text }
    $tr.Font.Name = $font
    $tr.Font.Size = $size
    $tr.Font.Bold = if ($bold) { -1 } else { 0 }
    $tr.Font.Fill.Solid()
    $tr.Font.Fill.ForeColor.RGB = $bgr
    $tr.ParagraphFormat.Alignment = $align
    return $tb
}

function Add-CapsulePill($slide, [string]$name, [string]$text, [single]$x, [single]$y, [single]$w, [single]$h, [int]$bgBgr, [single]$bgTrans, [int]$fgBgr, [single]$fontSize, [string]$fontName = $DisplayFont) {
    $sh = Find-Shape $slide $name
    if ($null -eq $sh) {
        $sh = $slide.Shapes.AddShape(5, $x, $y, $w, $h)
        $sh.Name = $name
        $sh.Adjustments.Item(1) = [single]0.5
    } else {
        $sh.Left = $x; $sh.Top = $y; $sh.Width = $w; $sh.Height = $h
    }
    $sh.Fill.Solid()
    $sh.Fill.ForeColor.RGB = $bgBgr
    $sh.Fill.Transparency = $bgTrans
    $sh.Line.Visible = 0
    $tf = $sh.TextFrame2
    $tf.WordWrap = 0
    $tf.AutoSize = 0
    $tf.VerticalAnchor = 3
    $tf.MarginLeft = [single]8
    $tf.MarginRight = [single]8
    $tr = $tf.TextRange
    $tr.Text = [string]$text
    $tr.Font.Name = $fontName
    $tr.Font.Size = $fontSize
    $tr.Font.Bold = -1
    $tr.Font.Fill.Solid()
    $tr.Font.Fill.ForeColor.RGB = $fgBgr
    $tr.ParagraphFormat.Alignment = 2
    return $sh
}

function Add-TwoToneTitle($slide, [string]$name, [string]$part1, [string]$part2, [single]$x, [single]$y, [single]$w, [single]$h, [single]$size, [int]$bgr1, [int]$bgr2, [int]$align = 2) {
    $tb = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $tb.Name = $name
    $tf = $tb.TextFrame2
    $tf.WordWrap = -1
    $tf.AutoSize = 0
    $tf.VerticalAnchor = 3
    $tr = $tf.TextRange
    $tr.Text = "$part1$part2"
    if ($part1.Length -gt 0) {
        $r1 = $tr.Characters(1, $part1.Length)
        $r1.Font.Name = $DisplayFont; $r1.Font.Size = $size; $r1.Font.Bold = -1
        $r1.Font.Fill.Solid(); $r1.Font.Fill.ForeColor.RGB = $bgr1
    }
    if ($part2.Length -gt 0) {
        $r2 = $tr.Characters($part1.Length + 1, $part2.Length)
        $r2.Font.Name = $DisplayFont; $r2.Font.Size = $size; $r2.Font.Bold = -1
        $r2.Font.Fill.Solid(); $r2.Font.Fill.ForeColor.RGB = $bgr2
    }
    $tr.ParagraphFormat.Alignment = $align
    return $tb
}

function Add-DarkKpiBox($slide, [string]$name, [string]$val, [string]$lbl, [single]$x, [single]$y, [single]$w, [single]$h, [int]$darkBgBgr, [int]$accentBgr, [int]$textLightBgr) {
    $sh = Add-Rect $slide $name $x $y $w $h $darkBgBgr ([single]0.16) $accentBgr ([single]1.2) ([single]0.14)
    $sh.Line.Transparency = [single]0.30
    $tf = $sh.TextFrame2; $tf.WordWrap = -1; $tf.AutoSize = 0; $tf.VerticalAnchor = 3
    $tf.MarginLeft = [single]10; $tf.MarginRight = [single]10
    $tr = $tf.TextRange
    $tr.Text = "$val`r$lbl"
    $p1 = $tr.Paragraphs(1); $p1.Font.Name = $DisplayFont; $p1.Font.Size = [single]24; $p1.Font.Bold = -1
    $p1.Font.Fill.Solid(); $p1.Font.Fill.ForeColor.RGB = $accentBgr; $p1.ParagraphFormat.Alignment = 2
    $p2 = $tr.Paragraphs(2); $p2.Font.Name = $SemiFont; $p2.Font.Size = [single]9.5; $p2.Font.Bold = 0
    $p2.Font.Fill.Solid(); $p2.Font.Fill.ForeColor.RGB = $textLightBgr; $p2.ParagraphFormat.Alignment = 2
    return $sh
}

function Ensure-Model3D($slide, [string]$name, [string]$glbPath, [single]$x, [single]$y, [single]$w, [single]$h, [single]$rx, [single]$ry, [single]$rz) {
    $m = Find-Shape $slide $name
    if ($null -eq $m) {
        $m = $slide.Shapes.Add3DModel($glbPath, 0, 1, $x, $y, $w, $h)
        $m.Name = $name
        $m.LockAspectRatio = -1
        $m.Model3D.AutoFit = 0
    } else {
        $m.Left = $x; $m.Top = $y; $m.Width = $w; $m.Height = $h
    }
    $m.Model3D.RotationX = $rx
    $m.Model3D.RotationY = $ry
    $m.Model3D.RotationZ = $rz
    $m.ZOrder(0)
    return $m
}

$cSurface   = HexToBgr $spec.palette.surface "#141824"
$cSurfAlt   = HexToBgr $spec.palette.surface_alt "#1C2233"
$cBorder    = HexToBgr $spec.palette.border "#2E384D"
$cAcc1      = HexToBgr $spec.palette.accent1 "#2563EB"
$cAcc1Dark  = HexToBgr $spec.palette.accent1_dark "#1D4ED8"
$cAcc2      = HexToBgr $spec.palette.accent2 "#10B981"
$cAcc3      = HexToBgr $spec.palette.accent3 "#F59E0B"
$cWhite     = HexToBgr "#FFFFFF"
$cLightText = HexToBgr "#E2E8F0"
$cMuted     = HexToBgr "#94A3B8"
$cIvoryInk  = HexToBgr "#111827"
$cInkMid    = HexToBgr "#1E293B"
$cMutedDark = HexToBgr "#475569"
$cObsidian  = HexToBgr $spec.palette.obsidian "#090A0D"

$cardAdj     = [single]0.08
$heroArch    = [string]$spec.hero_3d.archetype
$isPhoneHero = ($heroArch -eq "iphone" -or $heroArch -eq "smartphone_pro")
$heroW       = if ($isPhoneHero) { [single]224 } else { [single]296 }
$heroH       = if ($isPhoneHero) { [single]448 } else { [single]320 }

$iphoneGlb  = if ($spec.requirements.iphone_glb) { [string]$spec.requirements.iphone_glb } else { Join-Path $AssetsDir "device_iphone.glb" }
$macbookGlb = if ($spec.requirements.macbook_glb) { [string]$spec.requirements.macbook_glb } else { Join-Path $AssetsDir "device_macbook.glb" }

if (!(Test-Path $ExportDir)) { New-Item -ItemType Directory -Force -Path $ExportDir | Out-Null }
$outParent = Split-Path -Parent $OutputPptx
if ($outParent -and !(Test-Path $outParent)) { New-Item -ItemType Directory -Force -Path $outParent | Out-Null }

Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $null

try {
    $pres = $ppt.Presentations.Add(0) # msoFalse headless
    $pres.PageSetup.SlideWidth = [single]960
    $pres.PageSetup.SlideHeight = [single]540

    $slides = @($spec.slides)
    $prevSlide = $null
    $prevType  = ""

    for ($si = 0; $si -lt $slides.Count; $si++) {
        $sl = $slides[$si]
        $idx = $si + 1
        $st = [string]$sl.type
        $nextType = if (($si + 1) -lt $slides.Count) { [string]$slides[$si + 1].type } else { "" }
        Write-Host ("Building Slide {0}/{1} ({2} | 3D={3})..." -f $idx, $slides.Count, $st, $heroArch)

        switch ($st) {
            "cover_hero" {
                $style = if ($sl.style) { [string]$sl.style } else { "centered_clean" }
                $s = Start-MorphSlide $pres $null @()

                if ($style -eq "3d_staged") {
                    Add-Bg $s "$AssetsDir\bg_ivory_grain.png" | Out-Null
                    Add-Text $s "s1_top" $sl.overline 80 34 800 22 $SemiFont 9.5 $false $cMutedDark 2 3 | Out-Null
                    Add-Rect $s "s1_line" 410 62 140 ([single]2.0) $cAcc1Dark 0 -1 0 0 | Out-Null
                    Add-Pic $s "s1_logo" "$AssetsDir\logo_badge.png" 426 112 108 108 | Out-Null
                    Add-TwoToneTitle $s "s1_title" $spec.brand.title_part1 $spec.brand.title_part2 60 228 840 94 74 $cIvoryInk $cAcc1Dark 2 | Out-Null
                    Add-Text $s "s1_motto" $spec.brand.motto 120 326 720 30 $DisplayFont 15 $true $cInkMid 2 3 | Out-Null
                    Add-Text $s "s1_sub" $sl.subtitle 120 364 720 26 $BodyFont 12.5 $false $cMutedDark 2 3 | Out-Null
                    Add-Text $s "s1_cue" $sl.bottom_cue 160 492 640 20 $SemiFont 9 $false $cMutedDark 2 3 | Out-Null
                    Add-Pic $s "!!SplitLower" "$AssetsDir\split_lower.png" 0 565 960 265 | Out-Null
                    Add-Text $s "!!GiantWord" "PROJECT" 30 -230 900 190 $DisplayFont 152 $true $cAcc1Dark 2 3 | Out-Null
                    Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 324 585 312 68 | Out-Null
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 368 575 $heroW $heroH -25 180 -12 | Out-Null
                    Set-Morph $s ([single]0.5)
                }
                elseif ($style -eq "asymmetric_split") {
                    Add-Bg $s "$AssetsDir\bg_ivory_grain.png" | Out-Null
                    Add-Text $s "c_ovr" $sl.overline 52 40 460 20 $MonoFont 9.5 $true $cAcc1Dark 1 3 | Out-Null
                    Add-Rect $s "c_rule" 52 66 84 ([single]2.5) $cAcc1Dark 0 -1 0 0 | Out-Null
                    Add-Pic $s "c_logo" "$AssetsDir\logo_badge.png" 52 86 54 54 | Out-Null
                    Add-Text $s "c_brand" $spec.brand.name 118 96 380 34 $DisplayFont 20 $true $cIvoryInk 1 3 | Out-Null
                    $head = if ($sl.headline) { [string]$sl.headline } else { "$($spec.brand.title_part1) $($spec.brand.title_part2)" }
                    Add-Text $s "c_h1" $head 52 152 468 134 $DisplayFont 42 $true $cIvoryInk 1 1 | Out-Null
                    Add-Text $s "c_sub" $sl.subtitle 52 296 460 66 $BodyFont 13.5 $false $cInkMid 1 1 | Out-Null

                    $meta = @($sl.meta)
                    $mCols = @($cAcc1, $cAcc2)
                    for ($mi = 0; $mi -lt [Math]::Min(2, $meta.Count); $mi++) {
                        $mx = [single](52 + $mi * 236)
                        Add-DarkKpiBox $s "c_mk_$mi" $meta[$mi].value $meta[$mi].label $mx 396 220 84 $cObsidian $mCols[$mi] $cLightText | Out-Null
                    }

                    Add-Pic $s "c_vis" "$AssetsDir\s$idx`_cover_visual.png" 556 32 362 476 | Out-Null
                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 568 424 336 62; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 586 88 302 346 14 -26 3 | Out-Null
                    Set-Morph $s ([single]0.6)
                }
                elseif ($style -eq "bold_poster") {
                    Add-Bg $s "$AssetsDir\bg_finale.png" | Out-Null
                    Add-Rect $s "c_band" 0 0 960 10 $cAcc1 0 -1 0 0 | Out-Null
                    Add-Text $s "c_ovr" $sl.overline 52 38 520 20 $MonoFont 10 $true $cAcc1 1 3 | Out-Null
                    $head = if ($sl.headline) { [string]$sl.headline } else { "$($spec.brand.title_part1)`n$($spec.brand.title_part2)" }
                    Add-Text $s "c_h1" $head 52 84 520 210 $HeadingFont 56 $true $cWhite 1 1 | Out-Null
                    Add-Rect $s "c_bar" 52 312 136 5 $cAcc2 0 -1 0 0 | Out-Null
                    Add-Text $s "c_sub" $sl.subtitle 52 332 500 76 $BodyFont 14 $false $cLightText 1 1 | Out-Null
                    if ($sl.bottom_cue) {
                        Add-CapsulePill $s "c_cue" $sl.bottom_cue 52 436 330 36 $cAcc1 0 $cWhite 10.5 $MonoFont | Out-Null
                    }
                    Add-Pic $s "c_vis" "$AssetsDir\s$idx`_cover_visual.png" 588 36 330 468 | Out-Null
                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 584 422 336 64; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 602 74 304 358 12 -24 -4 | Out-Null
                    Set-Morph $s ([single]0.6)
                }
                else {
                    Add-Bg $s "$AssetsDir\bg_main.png" | Out-Null
                    Add-Text $s "c_ovr" $sl.overline 120 18 720 18 $MonoFont 9.5 $true $cAcc1 2 3 | Out-Null
                    Add-Pic $s "c_logo" "$AssetsDir\logo_badge.png" 52 28 56 56 | Out-Null
                    Add-TwoToneTitle $s "c_title" $spec.brand.title_part1 $spec.brand.title_part2 120 42 720 66 48 $cWhite $cAcc1 2 | Out-Null
                    Add-Text $s "c_sub" $sl.subtitle 140 112 680 34 $BodyFont 12.5 $false $cLightText 2 3 | Out-Null

                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 320 392 320 58; $shP.ZOrder(0)
                    $mW = if ($isPhoneHero) { [single]170 } else { [single]284 }
                    $mH = if ($isPhoneHero) { [single]262 } else { [single]252 }
                    $mX = [single](480 - ($mW / 2))
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb $mX 152 $mW $mH 14 -22 3 | Out-Null

                    $meta = @($sl.meta)
                    $mCols = @($cAcc1, $cAcc2, $cAcc3)
                    for ($mi = 0; $mi -lt [Math]::Min(3, $meta.Count); $mi++) {
                        $mx = [single](132 + $mi * 242)
                        Add-DarkKpiBox $s "c_mk_$mi" $meta[$mi].value $meta[$mi].label $mx 434 212 76 $cObsidian $mCols[$mi] $cLightText | Out-Null
                    }
                    Set-Morph $s ([single]0.6)
                }

                # Pre-stage off-canvas Morph actors for Slide 2 when Slide 2 uses horizon or depth-sandwich transitions
                if (($si + 1) -lt $slides.Count) {
                    $nextSl = $slides[$si + 1]
                    $nextGw = if ($nextSl.giant_word) { [string]$nextSl.giant_word } else { "PROJECT" }
                    if ($nextType -eq "split_horizon_3d" -and $style -ne "3d_staged") {
                        Add-Pic $s "!!SplitLower" "$AssetsDir\split_lower.png" 0 565 960 265 | Out-Null
                        Add-Text $s "!!GiantWord" $nextGw 30 -230 900 180 $DisplayFont 122 $true $cAcc1Dark 2 3 | Out-Null
                    } elseif ($nextType -eq "depth_sandwich_pan") {
                        Add-Text $s "!!pan_word" $nextGw 30 -230 900 196 $HeadingFont 136 $true $cWhite 2 3 | Out-Null
                    } elseif ($nextType -eq "macro_dolly_3d") {
                        for ($mi = 0; $mi -lt 3; $mi++) {
                            Add-Pic $s "!!md_card_$mi" "$AssetsDir\card_spec_$($mi + 1).png" (-460 - $mi * 140) (142 + $mi * 116) 384 96 | Out-Null
                        }
                    }
                }
            }

            "portal_flythrough_3d" {
                $stage = if ($sl.stage) { [string]$sl.stage } else { "exterior" }
                $keep = @("!!PhoneShadow", "!!Phone3D")
                if ($prevType -eq "portal_flythrough_3d") {
                    $keep += @("!!portal_bg", "!!portal_inner", "!!portal_outer", "!!pt_card_next")
                }
                $s = Start-MorphSlide $pres $prevSlide $keep

                if ($stage -eq "interior") {
                    $pbg = Add-Pic $s "!!portal_bg" "$AssetsDir\portal_bg.png" -64 -36 1088 612; $pbg.ZOrder(1)
                    $pout = Add-Pic $s "!!portal_outer" "$AssetsDir\portal_outer.png" -1728 -972 4416 2484; $pout.ZOrder(0)
                    $pin = Add-Pic $s "!!portal_inner" "$AssetsDir\portal_inner.png" 0 0 960 540; $pin.ZOrder(0)

                    $pcard = Add-Pic $s "!!pt_card_next" "$AssetsDir\card_bento_wide.png" 64 108 430 324; $pcard.ZOrder(0)
                    Add-Text $s "pt_ovr" $sl.overline 88 126 382 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "pt_h1" $sl.headline 88 150 382 68 $DisplayFont 24 $true $cWhite 1 1 | Out-Null
                    $pBody = if ($sl.body) { [string]$sl.body } else { [string]$sl.subtitle }
                    Add-Text $s "pt_body" $pBody 88 224 382 92 $BodyFont 11.5 $false $cLightText 1 1 | Out-Null
                    $kpis = @($sl.kpis); $kCols = @($cAcc1, $cAcc2)
                    for ($ki = 0; $ki -lt [Math]::Min(2, $kpis.Count); $ki++) {
                        Add-DarkKpiBox $s "pt_k_$ki" $kpis[$ki].value $kpis[$ki].label (88 + $ki * 194) 330 182 76 $cObsidian $kCols[$ki] $cLightText | Out-Null
                    }

                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 526 408 348 64; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 544 72 316 352 12 -20 2 | Out-Null
                    Set-Morph $s ([single]2.1)
                } elseif ($stage -eq "horizon") {
                    $pbg = Add-Pic $s "!!portal_bg" "$AssetsDir\portal_bg.png" -120 -68 1200 675; $pbg.ZOrder(1)
                    $pout = Add-Pic $s "!!portal_outer" "$AssetsDir\portal_outer.png" -1728 -972 4416 2484; $pout.ZOrder(0)
                    $pin = Add-Pic $s "!!portal_inner" "$AssetsDir\portal_inner.png" -1536 -864 4032 2268; $pin.ZOrder(0)

                    Add-Text $s "pt_ovr" $sl.overline 52 46 460 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    $ph1 = Add-Text $s "pt_h1" $sl.headline 52 72 460 96 $DisplayFont 32 $true $cWhite 1 1
                    Add-TextShadow $ph1
                    $pBody = if ($sl.body) { [string]$sl.body } else { [string]$sl.subtitle }
                    Add-Text $s "pt_body" $pBody 52 180 440 110 $BodyFont 12.5 $false $cLightText 1 1 | Out-Null
                    $kpis = @($sl.kpis); $kCols = @($cAcc1, $cAcc2, $cAcc3)
                    for ($ki = 0; $ki -lt [Math]::Min(3, $kpis.Count); $ki++) {
                        Add-DarkKpiBox $s "pt_k_$ki" $kpis[$ki].value $kpis[$ki].label (52 + $ki * 158) 396 146 78 $cObsidian $kCols[$ki] $cLightText | Out-Null
                    }

                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 522 424 388 68; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 540 58 356 384 14 24 -3 | Out-Null
                    Set-Morph $s ([single]2.0)
                } else {
                    $pbg = Add-Pic $s "!!portal_bg" "$AssetsDir\portal_bg.png" 0 0 960 540; $pbg.ZOrder(1)
                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 380 344 200 44; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 392 146 176 214 16 -28 -4 | Out-Null
                    $pin = Add-Pic $s "!!portal_inner" "$AssetsDir\portal_inner.png" 272 102 416 336; $pin.ZOrder(0)
                    $pout = Add-Pic $s "!!portal_outer" "$AssetsDir\portal_outer.png" 0 0 960 540; $pout.ZOrder(0)

                    Add-Text $s "pt_ovr" $sl.overline 48 20 864 18 $MonoFont 9.5 $true $cAcc1 2 3 | Out-Null
                    $ph1 = Add-Text $s "pt_h1" $sl.headline 48 40 864 42 $DisplayFont 25 $true $cWhite 2 3
                    Add-TextShadow $ph1
                    $meta = @($sl.meta); $mCols = @($cAcc1, $cAcc2)
                    for ($mi = 0; $mi -lt [Math]::Min(2, $meta.Count); $mi++) {
                        $mx = if ($mi -eq 0) { [single]34 } else { [single]720 }
                        Add-DarkKpiBox $s "pt_m_$mi" $meta[$mi].value $meta[$mi].label $mx 204 206 112 $cObsidian $mCols[$mi] $cLightText | Out-Null
                    }
                    if ($sl.bottom_cue) {
                        $cuePill = Add-CapsulePill $s "pt_cue" $sl.bottom_cue 290 480 380 34 $cObsidian ([single]0.12) $cAcc1 10 $MonoFont
                        $cuePill.Line.Visible = -1; $cuePill.Line.ForeColor.RGB = $cBorder; $cuePill.Line.Weight = [single]1.2
                    }
                    if ($nextType -eq "portal_flythrough_3d") {
                        Add-Pic $s "!!pt_card_next" "$AssetsDir\card_bento_wide.png" -520 108 430 324 | Out-Null
                    }
                    Set-Morph $s ([single]2.0)
                }
            }

            "vertical_elevator_3d" {
                $stop = if ($sl.elevator_stop) { [string]$sl.elevator_stop } else { "upper_sky" }
                $gw = if ($sl.giant_word) { [string]$sl.giant_word } else { "ORBITAL" }
                $keep = @("!!PhoneShadow", "!!Phone3D")
                if ($prevType -eq "vertical_elevator_3d") {
                    $keep += @("!!elev_bg", "!!elev_word", "!!elev_fg", "!!ev_rcard_0", "!!ev_rcard_1", "!!ev_rcard_2")
                }
                $s = Start-MorphSlide $pres $prevSlide $keep

                if ($stop -eq "lower_ground") {
                    $ebg = Add-Pic $s "!!elev_bg" "$AssetsDir\elev_bg.png" 0 -520 960 1080; $ebg.ZOrder(1)
                    $ew = Add-Text $s "!!elev_word" $gw 40 290 880 170 $HeadingFont 132 $true $cWhite 2 3
                    Add-TextShadow $ew; $ew.ZOrder(0)
                    $efg = Add-Pic $s "!!elev_fg" "$AssetsDir\elev_fg.png" 0 -520 960 1080; $efg.ZOrder(0)

                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 46 418 356 62; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 62 74 328 362 16 28 -3 | Out-Null

                    Add-Text $s "ev_ovr2" $sl.overline 440 34 474 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    $eh2 = Add-Text $s "ev_h2" $sl.headline 440 56 474 68 $DisplayFont 25 $true $cWhite 1 1
                    Add-TextShadow $eh2

                    $cards = if ($sl.cards) { @($sl.cards) } else { @($sl.kpis) }
                    $rCols = @($cAcc1, $cAcc2, $cAcc3)
                    for ($ci = 0; $ci -lt [Math]::Min(3, $cards.Count); $ci++) {
                        $cy = [single](142 + $ci * 120); $cIdx = $ci + 1; $crd = $cards[$ci]
                        $rc = Add-Pic $s "!!ev_rcard_$ci" "$AssetsDir\card_spec_$cIdx.png" 440 $cy 474 104; $rc.ZOrder(0)
                        $met = if ($crd.value) { [string]$crd.value } else { [string]$crd.metric }
                        $ttl = if ($crd.title) { [string]$crd.title } else { [string]$crd.label }
                        Add-Text $s "ev_rmet_$ci" $met 462 ($cy + 16) 128 68 $DisplayFont 28 $true $rCols[$ci] 2 3 | Out-Null
                        Add-Text $s "ev_rttl_$ci" $ttl 598 ($cy + 16) 294 26 $DisplayFont 14 $true $cWhite 1 3 | Out-Null
                        Add-Text $s "ev_rdsc_$ci" $crd.desc 598 ($cy + 46) 294 44 $BodyFont 11 $false $cLightText 1 1 | Out-Null
                    }
                    Set-Morph $s ([single]2.2)
                } else {
                    $ebg = Add-Pic $s "!!elev_bg" "$AssetsDir\elev_bg.png" 0 0 960 1080; $ebg.ZOrder(1)
                    $ew = Add-Text $s "!!elev_word" $gw 40 94 880 156 $HeadingFont 132 $true $cWhite 2 3
                    Add-TextShadow $ew; $ew.ZOrder(0)
                    $efg = Add-Pic $s "!!elev_fg" "$AssetsDir\elev_fg.png" 0 0 960 1080; $efg.ZOrder(0)

                    $ep = Add-Pic $s "ev_panel" "$AssetsDir\card_bento_wide.png" 46 250 468 242; $ep.ZOrder(0)
                    Add-Text $s "ev_ovr" $sl.overline 72 268 416 18 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "ev_h1" $sl.headline 72 292 416 62 $DisplayFont 22 $true $cWhite 1 1 | Out-Null
                    Add-Text $s "ev_body" $sl.body 72 358 416 76 $BodyFont 11.5 $false $cLightText 1 1 | Out-Null
                    if ($sl.badge) { Add-CapsulePill $s "ev_badge" $sl.badge 72 444 198 28 $cAcc1 0 $cWhite 9.5 $MonoFont | Out-Null }

                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 548 412 356 62; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 564 68 326 360 10 -22 3 | Out-Null

                    if ($nextType -eq "vertical_elevator_3d") {
                        for ($ci = 0; $ci -lt 3; $ci++) {
                            $ox = [single](1080 + $ci * 160); $cy = [single](142 + $ci * 120); $cIdx = $ci + 1
                            Add-Pic $s "!!ev_rcard_$ci" "$AssetsDir\card_spec_$cIdx.png" $ox $cy 474 104 | Out-Null
                        }
                    }
                    Set-Morph $s ([single]2.1)
                }
            }

            "depth_sandwich_pan" {
                $keep = @("!!PhoneShadow", "!!Phone3D", "!!pan_word")
                if ($prevType -eq "depth_sandwich_pan") { $keep += @("!!pan_bg", "!!pan_fg") }
                $s = Start-MorphSlide $pres $prevSlide $keep

                $panState = if ($sl.pan_state) { [string]$sl.pan_state } else { "wide" }
                if ($panState -eq "zoom_right") {
                    $bx = [single]-240; $by = [single]-96; $bw = [single]1344; $bh = [single]756
                    $mx = [single]528;  $my = [single]46;  $mw = [single]366;  $mh = [single]396; $rx = [single]14; $ry = [single]-28; $rz = [single]3
                    $wx = [single]44;   $wy = [single]96;  $ww = [single]660;  $wh = [single]190; $wSize = [single]116; $wAlign = 1
                    $capX = [single]46; $capY = [single]346; $capW = [single]448
                } elseif ($panState -eq "zoom_left") {
                    $bx = [single]-120; $by = [single]-108; $bw = [single]1344; $bh = [single]756
                    $mx = [single]58;   $my = [single]46;   $mw = [single]366;  $mh = [single]396; $rx = [single]14; $ry = [single]28;  $rz = [single]-3
                    $wx = [single]260;  $wy = [single]96;   $ww = [single]650;  $wh = [single]190; $wSize = [single]116; $wAlign = 3
                    $capX = [single]466; $capY = [single]346; $capW = [single]448
                } else {
                    $bx = [single]0;    $by = [single]0;   $bw = [single]960; $bh = [single]540
                    $mx = [single]318;  $my = [single]46;  $mw = [single]324; $mh = [single]356; $rx = [single]10; $ry = [single]-14; $rz = [single]2
                    $wx = [single]30;   $wy = [single]88;  $ww = [single]900; $wh = [single]200; $wSize = [single]142; $wAlign = 2
                    $capX = [single]46; $capY = [single]384; $capW = [single]416
                }

                $pbg = Add-Pic $s "!!pan_bg" "$AssetsDir\s$idx`_pan_bg.png" $bx $by $bw $bh; $pbg.ZOrder(1)
                $pw  = Add-Text $s "!!pan_word" ([string]$sl.giant_word) $wx $wy $ww $wh $HeadingFont $wSize $true $cWhite $wAlign 3; Add-TextShadow $pw; $pw.ZOrder(0)
                $fgPath = "$AssetsDir\s$idx`_pan_fg.png"
                if (Test-Path $fgPath) { $pfg = Add-Pic $s "!!pan_fg" $fgPath $bx $by $bw $bh; $pfg.ZOrder(0) }

                $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" ($mx - 12) ($my + $mh - 12) ($mw + 24) 64; $shP.ZOrder(0)
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb $mx $my $mw $mh $rx $ry $rz | Out-Null

                Add-CapsulePill $s "pan_ovr" $sl.overline 46 28 336 28 $cObsidian ([single]0.15) $cAcc1 9.5 $MonoFont | Out-Null
                Add-Pic $s "pan_cap_bg" "$AssetsDir\card_bento_wide.png" $capX $capY $capW 124 | Out-Null
                Add-Text $s "pan_h" $sl.headline ($capX + 24) ($capY + 14) ($capW - 40) 34 $DisplayFont 18 $true $cWhite 1 3 | Out-Null
                Add-Text $s "pan_b" $sl.caption ($capX + 24) ($capY + 52) ($capW - 40) 56 $BodyFont 11.5 $false $cLightText 1 1 | Out-Null
                Set-Morph $s ([single]2.1)
            }

            "macro_dolly_3d" {
                $keep = @("!!PhoneShadow", "!!Phone3D", "!!md_card_0", "!!md_card_1", "!!md_card_2")
                if ($prevType -eq "split_horizon_3d") { $keep += @("!!SplitLower", "!!GiantWord") }
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Bg $s "$AssetsDir\bg_alt.png" | Out-Null
                $low = Find-Shape $s "!!SplitLower"; if ($null -ne $low) { $low.Top = [single]585 }
                $gw  = Find-Shape $s "!!GiantWord";  if ($null -ne $gw)  { $gw.Top  = [single]-250 }

                Add-Text $s "md_ovr" $sl.overline 46 22 520 18 $MonoFont 10 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "md_h1" $sl.headline 46 46 470 76 $DisplayFont 25 $true $cWhite 1 1 | Out-Null

                $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_laptop.png" 486 436 440 76; $shP.ZOrder(0)
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb 504 42 416 428 22 -38 -4 | Out-Null

                $callouts = @($sl.callouts)
                $cYs = @(142, 258, 374); $cCols = @($cAcc1, $cAcc2, $cAcc3)
                for ($ci = 0; $ci -lt [Math]::Min(3, $callouts.Count); $ci++) {
                    $cy = [single]$cYs[$ci]; $cIdx = $ci + 1; $ac = $cCols[$ci]
                    $mc = Add-Pic $s "!!md_card_$ci" "$AssetsDir\card_spec_$cIdx.png" 46 $cy 384 96; $mc.ZOrder(0)
                    Add-Text $s "md_ttl_$ci" $callouts[$ci].title 68 ($cy + 14) 344 22 $DisplayFont 13 $true $ac 1 3 | Out-Null
                    Add-Text $s "md_dsc_$ci" $callouts[$ci].desc 68 ($cy + 40) 344 46 $BodyFont 11 $false $cLightText 1 1 | Out-Null
                    Add-Rect $s "md_line_$ci" 430 ($cy + 46) 96 ([single]2.0) $ac 0 -1 0 0 | Out-Null
                    $dot = $s.Shapes.AddShape(9, 522, ($cy + 42), 10, 10)
                    $dot.Name = "md_dot_$ci"; $dot.Fill.Solid(); $dot.Fill.ForeColor.RGB = $ac; $dot.Line.Visible = 0
                }
                Set-Morph $s ([single]2.1)
            }

            "timeline_morph_step" {
                $stages = @($sl.stages)
                $nStages = [Math]::Max(1, [Math]::Min(4, $stages.Count))
                $actIdx = [Math]::Max(0, [Math]::Min($nStages - 1, [int]$sl.active_index))

                $keep = @("!!PhoneShadow", "!!Phone3D")
                if ($prevType -eq "timeline_morph_step") {
                    $keep += @("bg_slide", "!!tl_rail", "!!tl_progress", "!!tl_pointer", "!!tl_drawer")
                    for ($ti = 0; $ti -lt $nStages; $ti++) { $keep += "!!tl_node_$ti" }
                }
                $s = Start-MorphSlide $pres $prevSlide $keep
                if ($null -eq (Find-Shape $s "bg_slide")) { Add-Bg $s "$AssetsDir\bg_main.png" | Out-Null }
                else { $s.Shapes.Item("bg_slide").ZOrder(1) }

                Add-Text $s "tl_ovr" $sl.overline 52 24 620 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "tl_h1" $sl.headline 52 46 760 56 $DisplayFont 25 $true $cWhite 1 1 | Out-Null

                $railX0 = [single]76; $railSpan = [single]808
                Add-Rect $s "!!tl_rail" $railX0 144 $railSpan 4 $cBorder 0 -1 0 0 | Out-Null

                $nodeCenters = @()
                for ($ti = 0; $ti -lt $nStages; $ti++) {
                    if ($nStages -eq 1) { $nodeCenters += [single]480 }
                    else { $nodeCenters += [single]($railX0 + 70 + (($railSpan - 140) * ($ti / [single]($nStages - 1)))) }
                }
                $actCx = $nodeCenters[$actIdx]
                $progW = [Math]::Max([single]24, [single]($actCx - $railX0))
                Add-Rect $s "!!tl_progress" $railX0 144 $progW 4 $cAcc1 0 -1 0 0 | Out-Null

                for ($ti = 0; $ti -lt $nStages; $ti++) {
                    $stg = $stages[$ti]
                    $cx = $nodeCenters[$ti]
                    if ($ti -eq $actIdx) {
                        $nw = [single]172; $nh = [single]36; $ny = [single]128; $bgFill = $cAcc1; $fgTxt = $cWhite; $bCol = $cAcc1
                    } elseif ($ti -lt $actIdx) {
                        $nw = [single]148; $nh = [single]30; $ny = [single]131; $bgFill = $cSurface; $fgTxt = $cAcc1; $bCol = $cAcc1
                    } else {
                        $nw = [single]148; $nh = [single]30; $ny = [single]131; $bgFill = $cSurface; $fgTxt = $cMuted; $bCol = $cBorder
                    }
                    $nx = [Math]::Max([single]48, [Math]::Min([single](912 - $nw), [single]($cx - $nw / 2)))
                    $pill = Add-CapsulePill $s "!!tl_node_$ti" "$($stg.code)  $($stg.label)" $nx $ny $nw $nh $bgFill 0 $fgTxt 9.5 $MonoFont
                    if ($ti -ne $actIdx) { $pill.Line.Visible = -1; $pill.Line.ForeColor.RGB = $bCol; $pill.Line.Weight = [single]1.2 }
                    $pill.ZOrder(0)
                }

                $ptrX = [Math]::Max([single]326, [Math]::Min([single]836, [single]($actCx - 36)))
                Add-Rect $s "!!tl_pointer" $ptrX 192 72 4 $cAcc1 0 -1 0 0 | Out-Null

                $rotYs = @(22, -24, 38, -16); $rotXs = @(8, 16, 12, 6)
                $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 32 448 276 54; $shP.ZOrder(0)
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb 44 192 252 276 ([single]$rotXs[$actIdx % 4]) ([single]$rotYs[$actIdx % 4]) 2 | Out-Null

                $curStage = $stages[$actIdx]
                Add-Pic $s "!!tl_drawer" "$AssetsDir\card_timeline_drawer.png" 316 204 596 292 | Out-Null
                Add-Text $s "tl_code" "STAGE  $($curStage.code)" 344 226 176 20 $MonoFont 10 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "tl_metric" $curStage.metric 344 252 176 76 $DisplayFont 44 $true $cWhite 1 3 | Out-Null
                Add-Text $s "tl_mlbl" $curStage.metric_label 344 334 176 36 $SemiFont 11 $false $cLightText 1 1 | Out-Null
                if ($curStage.tag) { Add-CapsulePill $s "tl_tag" $curStage.tag 344 422 164 30 $cSurfAlt 0 $cAcc1 9 $MonoFont | Out-Null }

                Add-Rect $s "tl_vdiv" 526 228 ([single]1.0) 244 $cBorder 0 -1 0 0 | Out-Null
                Add-Text $s "tl_sh1" $curStage.title 544 226 346 52 $DisplayFont 18.5 $true $cWhite 1 1 | Out-Null
                Add-Text $s "tl_sum" $curStage.summary 544 284 346 58 $BodyFont 11.5 $false $cLightText 1 1 | Out-Null

                $bullets = @($curStage.bullets)
                for ($bi = 0; $bi -lt [Math]::Min(3, $bullets.Count); $bi++) {
                    $by = 354 + $bi * 40
                    Add-Rect $s "tl_bdot_$bi" 546 ($by + 9) 7 7 $cAcc1 0 -1 0 0 | Out-Null
                    Add-Text $s "tl_btxt_$bi" $bullets[$bi] 564 $by 328 32 $SemiFont 11 $false $cWhite 1 3 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "kpi_bento" {
                $keep = @("!!PhoneShadow", "!!Phone3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Bg $s "$AssetsDir\bg_main.png" | Out-Null
                Add-Text $s "kb_ovr" $sl.overline 48 24 620 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "kb_h1" $sl.headline 48 46 680 54 $DisplayFont 25 $true $cWhite 1 1 | Out-Null

                $hero = $sl.hero_tile
                $kbLayout = if ($sl.layout) { [string]$sl.layout } else { "left_anchor" }

                if ($kbLayout -eq "top_hero_strip") {
                    Add-Pic $s "kb_hero" "$AssetsDir\card_bento_strip.png" 46 106 588 154 | Out-Null
                    Add-Text $s "kb_htag" $hero.label 74 120 180 18 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "kb_hval" $hero.value 74 144 180 84 $DisplayFont 46 $true $cWhite 1 1 | Out-Null
                    Add-Rect $s "kb_vdiv" 262 128 ([single]1.0) 108 $cBorder 0 -1 0 0 | Out-Null
                    Add-Text $s "kb_hsub" $hero.title 280 124 334 30 $DisplayFont 16.5 $true $cWhite 1 3 | Out-Null
                    Add-Text $s "kb_hdsc" $hero.desc 280 158 334 62 $BodyFont 11.5 $false $cLightText 1 1 | Out-Null

                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 654 234 260 44; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 668 42 232 214 14 -28 3 | Out-Null

                    $bTiles = @($sl.tiles)
                    if ($bTiles.Count -lt 3 -and $null -ne $sl.wide_tile) {
                        $wt = $sl.wide_tile
                        $bTiles += @{ label = $wt.overline; value = if ($wt.value) { $wt.value } else { "99.9%" }; desc = $wt.title }
                    }
                    $bCols = @($cAcc1, $cAcc2, $cAcc3)
                    for ($ti = 0; $ti -lt [Math]::Min(3, $bTiles.Count); $ti++) {
                        $t = $bTiles[$ti]; $tx = [single](46 + $ti * 296); $ac = $bCols[$ti % 3]; $tIdx = $ti + 1
                        Add-Pic $s "kb_tbg_$ti" "$AssetsDir\card_bento_tile_$tIdx.png" $tx 276 276 222 | Out-Null
                        Add-Text $s "kb_tlbl_$ti" $t.label ($tx + 22) 298 232 20 $MonoFont 9.5 $true $ac 1 3 | Out-Null
                        Add-Text $s "kb_tval_$ti" $t.value ($tx + 22) 324 232 62 $DisplayFont 38 $true $cWhite 1 3 | Out-Null
                        Add-Text $s "kb_tdsc_$ti" $t.desc ($tx + 22) 396 232 82 $BodyFont 11.5 $false $cLightText 1 1 | Out-Null
                    }
                } else {
                    Add-Pic $s "kb_hero" "$AssetsDir\card_bento_hero.png" 46 110 284 390 | Out-Null
                    Add-Text $s "kb_htag" $hero.label 70 134 236 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "kb_hval" $hero.value 70 162 236 82 $DisplayFont 50 $true $cWhite 1 3 | Out-Null
                    Add-Text $s "kb_hsub" $hero.title 70 250 236 30 $DisplayFont 15 $true $cWhite 1 3 | Out-Null
                    Add-Text $s "kb_hdsc" $hero.desc 70 286 236 116 $BodyFont 12 $false $cLightText 1 1 | Out-Null
                    if ($hero.badge) { Add-CapsulePill $s "kb_hcap" $hero.badge 70 436 210 32 $cAcc1 0 $cWhite 9.5 $MonoFont | Out-Null }

                    $tiles = @($sl.tiles); $tCols = @($cAcc1, $cAcc2)
                    for ($ti = 0; $ti -lt [Math]::Min(2, $tiles.Count); $ti++) {
                        $t = $tiles[$ti]; $tx = [single](344 + $ti * 198); $ac = $tCols[$ti]; $tIdx = $ti + 1
                        Add-Pic $s "kb_tbg_$ti" "$AssetsDir\card_bento_tile_$tIdx.png" $tx 110 186 184 | Out-Null
                        Add-Text $s "kb_tlbl_$ti" $t.label ($tx + 18) 128 150 18 $MonoFont 9 $true $ac 1 3 | Out-Null
                        Add-Text $s "kb_tval_$ti" $t.value ($tx + 18) 150 150 54 $DisplayFont 32 $true $cWhite 1 3 | Out-Null
                        Add-Text $s "kb_tdsc_$ti" $t.desc ($tx + 18) 210 150 68 $BodyFont 10.5 $false $cLightText 1 1 | Out-Null
                    }

                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 738 264 184 42; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 744 64 176 218 14 -26 3 | Out-Null

                    $wide = $sl.wide_tile
                    Add-Pic $s "kb_wide" "$AssetsDir\card_bento_wide.png" 344 310 570 190 | Out-Null
                    Add-Text $s "kb_wovr" $wide.overline 368 326 520 18 $MonoFont 9 $true $cAcc2 1 3 | Out-Null
                    Add-Text $s "kb_wh1" $wide.title 368 348 520 30 $DisplayFont 16.5 $true $cWhite 1 3 | Out-Null
                    $cols = @($wide.columns)
                    for ($ci = 0; $ci -lt [Math]::Min(2, $cols.Count); $ci++) {
                        $cx = 368 + $ci * 264
                        Add-Text $s "kb_wcl_$ci" $cols[$ci].label $cx 390 246 20 $SemiFont 11.5 $true $cAcc1 1 3 | Out-Null
                        Add-Text $s "kb_wcd_$ci" $cols[$ci].desc $cx 414 246 66 $BodyFont 11 $false $cLightText 1 1 | Out-Null
                    }
                }
                Set-Morph $s ([single]1.8)
            }

            "chart_infographic" {
                $keep = @("!!PhoneShadow", "!!Phone3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Bg $s "$AssetsDir\bg_main.png" | Out-Null
                Add-Text $s "ch_ovr" $sl.overline 46 24 600 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "ch_h1" $sl.headline 46 46 600 54 $DisplayFont 25 $true $cWhite 1 1 | Out-Null
                Add-Pic $s "ch_img" "$AssetsDir\s$idx`_chart.png" 46 110 584 384 | Out-Null

                $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 656 244 248 44; $shP.ZOrder(0)
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb 672 38 220 222 14 -24 2 | Out-Null

                $hl = $sl.highlight
                Add-Pic $s "ch_hlbg" "$AssetsDir\card_spec_1.png" 648 278 268 104 | Out-Null
                Add-Text $s "ch_hlv" "$($hl.value)  $($hl.label)" 670 290 232 30 $DisplayFont 18 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "ch_hld" $hl.desc 670 324 232 48 $BodyFont 10.5 $false $cLightText 1 1 | Out-Null

                $insights = @($sl.insights)
                if ($insights.Count -gt 0) {
                    $in = $insights[0]
                    Add-Pic $s "ch_inbg_0" "$AssetsDir\card_spec_2.png" 648 394 268 100 | Out-Null
                    Add-Text $s "ch_int_0" $in.title 670 406 232 24 $DisplayFont 13 $true $cAcc2 1 3 | Out-Null
                    Add-Text $s "ch_ind_0" $in.desc 670 434 232 50 $BodyFont 10.5 $false $cLightText 1 1 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "architecture_diagram" {
                $keep = @("!!PhoneShadow", "!!Phone3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Bg $s "$AssetsDir\bg_alt.png" | Out-Null
                Add-Text $s "ar_ovr" $sl.overline 42 22 620 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "ar_h1" $sl.headline 42 46 630 52 $DisplayFont 24 $true $cWhite 1 1 | Out-Null
                Add-Pic $s "ar_img" "$AssetsDir\s$idx`_arch.png" 42 104 624 398 | Out-Null

                $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 686 244 236 42; $shP.ZOrder(0)
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb 698 36 212 222 22 -34 -3 | Out-Null

                $pillars = @($sl.pillars); $spYs = @(276, 392); $spCols = @($cAcc1, $cAcc2)
                for ($pi = 0; $pi -lt [Math]::Min(2, $pillars.Count); $pi++) {
                    $pl = $pillars[$pi]; $sy = [single]$spYs[$pi]; $ac = $spCols[$pi]; $cIdx = $pi + 1
                    Add-Pic $s "ar_cbg_$pi" "$AssetsDir\card_spec_$cIdx.png" 682 $sy 236 106 | Out-Null
                    Add-Text $s "ar_ctag_$pi" "$($pl.tag)  //  $($pl.title)" 702 ($sy + 12) 204 22 $DisplayFont 12 $true $ac 1 3 | Out-Null
                    Add-Text $s "ar_cdsc_$pi" $pl.desc 702 ($sy + 38) 204 58 $BodyFont 10.5 $false $cLightText 1 1 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "editorial_split" {
                $side = if ($sl.image_side) { [string]$sl.image_side } else { "right" }
                $keep = @("!!PhoneShadow", "!!Phone3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Bg $s "$AssetsDir\bg_main.png" | Out-Null

                if ($side -eq "left") { $imgX = [single]44; $txtX = [single]504; $mX = [single]228; $mRy = [single]24 }
                else { $txtX = [single]48; $imgX = [single]508; $mX = [single]476; $mRy = [single]-24 }

                Add-Pic $s "es_vis" "$AssetsDir\s$idx`_split_visual.png" $imgX 34 406 468 | Out-Null
                $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" ($mX - 10) 432 280 54; $shP.ZOrder(0)
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb $mX 118 260 326 12 $mRy 2 | Out-Null

                Add-Text $s "es_ovr" $sl.overline $txtX 44 416 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Rect $s "es_rule" $txtX 70 72 ([single]2.0) $cAcc1 0 -1 0 0 | Out-Null
                Add-Text $s "es_h1" $sl.headline $txtX 88 416 108 $DisplayFont 30 $true $cWhite 1 1 | Out-Null
                Add-Text $s "es_body" $sl.body $txtX 206 416 104 $BodyFont 13.5 $false $cLightText 1 1 | Out-Null

                $facts = @($sl.facts); $fCols = @($cAcc1, $cAcc2)
                for ($fi = 0; $fi -lt [Math]::Min(2, $facts.Count); $fi++) {
                    $fx = [single]($txtX + $fi * 212); $cIdx = $fi + 1
                    Add-Pic $s "es_fcard_$fi" "$AssetsDir\card_spec_$cIdx.png" $fx 342 198 136 | Out-Null
                    Add-Text $s "es_fv_$fi" $facts[$fi].value ($fx + 18) 354 166 40 $DisplayFont 26 $true $fCols[$fi] 1 3 | Out-Null
                    Add-Text $s "es_fl_$fi" $facts[$fi].label ($fx + 18) 396 166 18 $MonoFont 8.5 $true $cWhite 1 3 | Out-Null
                    Add-Text $s "es_fd_$fi" $facts[$fi].desc ($fx + 18) 416 166 52 $BodyFont 10 $false $cLightText 1 1 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "split_horizon_3d" {
                $keep = @("bg_slide", "!!SplitLower", "!!GiantWord", "!!PhoneShadow", "!!Phone3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                if ($null -eq (Find-Shape $s "bg_slide")) { Add-Bg $s "$AssetsDir\bg_ivory_grain.png" | Out-Null }
                else { $s.Shapes.Item("bg_slide").ZOrder(1) }

                $slLow = Add-Pic $s "!!SplitLower" "$AssetsDir\split_lower.png" 0 275 960 265; $slLow.ZOrder(0)
                $gwSize = if ($DisplayFont -eq "Georgia") { [single]122 } else { [single]148 }
                $gwTop  = if ($DisplayFont -eq "Georgia") { [single]38 } else { [single]14 }
                $gw = Add-Text $s "!!GiantWord" ([string]$sl.giant_word) 30 $gwTop 900 180 $DisplayFont $gwSize $true $cAcc1Dark 2 3; $gw.ZOrder(0)

                Add-Text $s "s2_pl" $sl.left_label 36 184 294 18 $DisplayFont 10 $true $cAcc1Dark 1 3 | Out-Null
                Add-Text $s "s2_pt" $sl.left_text 36 202 294 58 $SemiFont 12.5 $false $cInkMid 1 1 | Out-Null
                Add-Text $s "s2_sl" $sl.right_label 630 184 294 18 $DisplayFont 10 $true $cAcc1Dark 3 3 | Out-Null
                Add-Text $s "s2_st" $sl.right_text 630 202 294 58 $SemiFont 12.5 $false $cInkMid 3 1 | Out-Null

                $kpiCols = @($cAcc1, $cAcc3, $cAcc2, $cAcc2); $kpiXs = @(36, 192, 624, 780); $kpis = @($sl.kpis)
                for ($i = 0; $i -lt [Math]::Min(4, $kpis.Count); $i++) {
                    Add-DarkKpiBox $s "s2_k$i" $kpis[$i].value $kpis[$i].label $kpiXs[$i] 396 144 78 $cObsidian $kpiCols[$i] $cLightText | Out-Null
                }
                $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 324 472 312 62; $shP.ZOrder(0)
                $mLeft = [single](480 - ($heroW / 2))
                $mTop  = if ($isPhoneHero) { [single]44 } else { [single]84 }
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb $mLeft $mTop $heroW $heroH 0 0 0 | Out-Null
                Add-Pic $s "!!SplitUpper" "$AssetsDir\split_upper.png" 0 -340 960 318 | Out-Null
                Set-Morph $s ([single]2.2)
            }

            "orbit_steps_3d" {
                $variant = if ($sl.variant) { [string]$sl.variant } else { "obsidian_steps" }
                if ($variant -eq "upper_stats") {
                    $keep = @("bg_slide", "!!SplitLower", "!!GiantWord", "!!SplitUpper", "!!PhoneShadow", "!!Phone3D")
                    $s = Start-MorphSlide $pres $prevSlide $keep
                    if ($null -eq (Find-Shape $s "bg_slide")) { Add-Bg $s "$AssetsDir\bg_ivory_grain.png" | Out-Null }
                    else { $s.Shapes.Item("bg_slide").ZOrder(1) }
                    $low = Find-Shape $s "!!SplitLower"; if ($null -ne $low) { $low.Top = [single]580 }
                    $gw = Find-Shape $s "!!GiantWord"; if ($null -ne $gw) { $gw.Top = [single]-250 }

                    Add-Pic $s "!!SplitUpper" "$AssetsDir\split_upper.png" 0 0 960 318 | Out-Null
                    Add-Text $s "s3_ovr" $sl.overline 48 40 500 22 $DisplayFont 11 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "s3_h1" $sl.headline 48 68 500 100 $DisplayFont 34 $true $cWhite 1 1 | Out-Null
                    Add-Text $s "s3_body" $sl.body 48 176 490 92 $BodyFont 13 $false $cLightText 1 1 | Out-Null

                    $cXs = @(48, 224, 400); $stats = @($sl.stats)
                    for ($i = 0; $i -lt [Math]::Min(3, $stats.Count); $i++) {
                        $cx = $cXs[$i]
                        Add-Rect $s "s3_bar_$i" $cx 344 152 ([single]2.5) $cAcc1 0 -1 0 0 | Out-Null
                        Add-Text $s "s3_kv_$i" $stats[$i].value $cx 354 156 50 $DisplayFont 38 $true $cIvoryInk 1 1 | Out-Null
                        Add-Text $s "s3_kl_$i" $stats[$i].label $cx 404 156 20 $DisplayFont 10 $true $cAcc1Dark 1 3 | Out-Null
                        Add-Text $s "s3_kd_$i" $stats[$i].detail $cx 426 156 54 $BodyFont 10.5 $false $cMutedDark 1 1 | Out-Null
                    }
                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 576 468 288 60; $shP.ZOrder(0)
                    $mTop = if ($isPhoneHero) { [single]38 } else { [single]108 }
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 608 $mTop $heroW $heroH 6 -21 3 | Out-Null
                } else {
                    $keep = @("bg_slide", "!!PhoneShadow", "!!Phone3D")
                    $s = Start-MorphSlide $pres $prevSlide $keep
                    Add-Bg $s "$AssetsDir\bg_main.png" | Out-Null
                    $pSide = if ($sl.side) { [string]$sl.side } else { "left" }
                    $mTop = if ($isPhoneHero) { [single]38 } else { [single]108 }
                    if ($pSide -eq "right") {
                        $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 628 468 296 60; $shP.ZOrder(0)
                        Ensure-Model3D $s "!!Phone3D" $iphoneGlb 636 $mTop $heroW $heroH 6 -21 3 | Out-Null
                        $colX = [single]46
                    } else {
                        $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 36 468 296 60; $shP.ZOrder(0)
                        Ensure-Model3D $s "!!Phone3D" $iphoneGlb 54 $mTop $heroW $heroH 6 21 -3 | Out-Null
                        $colX = [single]366
                    }

                    Add-Text $s "s4_ovr" $sl.overline $colX 36 540 20 $MonoFont 10.5 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "s4_h1" $sl.headline $colX 56 550 72 $DisplayFont 29 $true $cWhite 1 3 | Out-Null

                    $stepYs = @(132, 222, 312, 402); $steps = @($sl.steps)
                    for ($i = 0; $i -lt [Math]::Min(4, $steps.Count); $i++) {
                        $sy = $stepYs[$i]
                        Add-Pic $s "s4_card_$i" "$AssetsDir\card_step.png" $colX $sy 546 76 | Out-Null
                        Add-Text $s "s4_num_$i" $steps[$i].num ($colX + 20) ($sy + 14) 56 48 $DisplayFont 26 $true $cAcc1 2 3 | Out-Null
                        Add-Text $s "s4_ttl_$i" $steps[$i].title ($colX + 86) ($sy + 13) 440 24 $DisplayFont 13.5 $true $cWhite 1 3 | Out-Null
                        Add-Text $s "s4_dsc_$i" $steps[$i].desc ($colX + 86) ($sy + 40) 445 24 $BodyFont 11.5 $false $cLightText 1 1 | Out-Null
                    }
                }
                Set-Morph $s ([single]2.2)
            }

            "spotlight_dock" {
                $dockIdx = if ($null -ne $sl.dock_index) { [int]$sl.dock_index } else { 0 }
                $dockXs = @(66, 282, 498, 714)
                $gliderXs = @(54, 270, 486, 702)
                $curGliderX = [single]$gliderXs[[Math]::Min(3, [Math]::Max(0, $dockIdx))]

                $keep = @("!!PhoneShadow", "!!Phone3D")
                if ($prevType -eq "spotlight_dock") {
                    $prevIdx = $idx - 1
                    $keep += @(
                        "!!sp_bg_$idx", "!!sp_dot_0", "!!sp_dot_1",
                        "!!sp_tag_$prevIdx", "!!sp_title_$prevIdx", "!!sp_sub_$prevIdx", "!!sp_body_$prevIdx",
                        "!!sp_tag_$idx", "!!sp_title_$idx", "!!sp_sub_$idx", "!!sp_body_$idx",
                        "!!GlassDock", "!!DockGlider",
                        "!!DockThumb1", "!!DockThumb2", "!!DockThumb3", "!!DockThumb4",
                        "!!DockLbl1", "!!DockLbl2", "!!DockLbl3", "!!DockLbl4",
                        "!!DockSub1", "!!DockSub2", "!!DockSub3", "!!DockSub4"
                    )
                }
                $s = Start-MorphSlide $pres $prevSlide $keep

                $curBg = Add-Pic $s "!!sp_bg_$idx" "$AssetsDir\s$idx`_bg.png" 0 0 960 540
                $curBg.ZOrder(1)

                # Orthogonal L-trajectory upward exit for previous spotlight text actors
                if ($prevType -eq "spotlight_dock") {
                    $prevIdx = $idx - 1
                    foreach ($pair in @(@("!!sp_tag_$prevIdx", -280), @("!!sp_title_$prevIdx", -240), @("!!sp_sub_$prevIdx", -180), @("!!sp_body_$prevIdx", -140))) {
                        $osh = Find-Shape $s $pair[0]
                        if ($null -ne $osh) { $osh.Top = [single]$pair[1] }
                    }
                }

                # Pre-stage incoming spotlight background with liquid vertical srcRect curtain wipe & staggered X text cascade
                if ($nextType -eq "spotlight_dock") {
                    $nextIdx = $idx + 1
                    $nbg = Add-Pic $s "!!sp_bg_$nextIdx" "$AssetsDir\s$nextIdx`_bg.png" 0 0 960 540
                    try {
                        $nbg.PictureFormat.CropTop = [single](-360)
                        $nbg.PictureFormat.CropBottom = [single]360
                        $nbg.Left = [single]0; $nbg.Top = [single]0; $nbg.Width = [single]960; $nbg.Height = [single]540
                    } catch {}
                    $nbg.ZOrder(1)
                    $nxSp = $slides[$si + 1].spotlight
                    Add-Text $s "!!sp_tag_$nextIdx" $nxSp.tag -460 48 540 20 $MonoFont 10 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "!!sp_title_$nextIdx" $nxSp.title -560 74 550 64 $DisplayFont 42 $true $cWhite 1 1 | Out-Null
                    Add-Text $s "!!sp_sub_$nextIdx" $nxSp.subtitle -680 142 560 26 $SemiFont 13 $false $cAcc3 1 3 | Out-Null
                    Add-Text $s "!!sp_body_$nextIdx" $nxSp.body -800 180 520 112 $BodyFont 13.5 $false $cLightText 1 1 | Out-Null
                }

                $sp = $sl.spotlight
                Add-Text $s "!!sp_tag_$idx" $sp.tag 50 48 540 20 $MonoFont 10 $true $cAcc1 1 3 | Out-Null
                $spTitle = Add-Text $s "!!sp_title_$idx" $sp.title 50 74 550 64 $DisplayFont 42 $true $cWhite 1 1
                Add-TextShadow $spTitle
                Add-Text $s "!!sp_sub_$idx" $sp.subtitle 50 142 560 26 $SemiFont 13 $false $cAcc3 1 3 | Out-Null
                Add-Text $s "!!sp_body_$idx" $sp.body 50 180 520 112 $BodyFont 13.5 $false $cLightText 1 1 | Out-Null

                if (-not $isPhoneHero -or [bool]$sl.show_3d) {
                    $rySp = if ($dockIdx % 2 -eq 0) { [single]-16 } else { [single]22 }
                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 644 304 268 46; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 652 38 252 276 10 $rySp 2 | Out-Null
                    if ($sp.badge) { Add-CapsulePill $s "sp_bdg" $sp.badge 686 328 184 30 $cObsidian ([single]0.12) $cAcc3 9.5 $MonoFont | Out-Null }
                } else {
                    Add-Pic $s "sp_port" "$AssetsDir\s$idx`_portrait.png" 656 36 244 282 | Out-Null
                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" -360 468 280 56
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb -340 60 $heroW $heroH 4 95 0 | Out-Null
                }

                $refractDock = "$AssetsDir\s$idx`_dock_refract.png"
                $dockImgPath = if (Test-Path $refractDock) { $refractDock } else { "$AssetsDir\dock_bg.png" }
                $gdSh = Find-Shape $s "!!GlassDock"
                if ($null -ne $gdSh) { $gdSh.Delete() }
                $gdSh = Add-Pic $s "!!GlassDock" $dockImgPath 40 386 880 128
                $gld = Add-Pic $s "!!DockGlider" "$AssetsDir\dock_glider.png" $curGliderX 396 208 108; $gld.ZOrder(0)

                $dockItems = if ($sl.dock) { @($sl.dock) } else { @($spec.dock) }
                for ($di = 0; $di -lt 4; $di++) {
                    $item = if ($di -lt $dockItems.Count) { $dockItems[$di] } else { @{ title = "Item $($di+1)"; subtitle = "" } }
                    $dnum = $di + 1; $dx = $dockXs[$di]
                    $th = Add-Pic $s "!!DockThumb$dnum" "$AssetsDir\thumb_$dnum.png" $dx 416 76 60; $th.ZOrder(0)
                    $tl = Add-Text $s "!!DockLbl$dnum" $item.title ($dx + 84) 418 110 24 $DisplayFont 13 $true $cWhite 1 3; $tl.ZOrder(0)
                    $ts = Add-Text $s "!!DockSub$dnum" $item.subtitle ($dx + 84) 444 110 22 $BodyFont 11 $false $cMuted 1 1; $ts.ZOrder(0)
                }

                for ($pi = 0; $pi -lt 2; $pi++) {
                    $isAct = ($pi -eq ($dockIdx % 2))
                    $dy = [single](222 + $pi * 24)
                    $dw = if ($isAct) { [single]14 } else { [single]7 }
                    $dxp = if ($isAct) { [single]930 } else { [single]933.5 }
                    $dot = Find-Shape $s "!!sp_dot_$pi"
                    if ($null -eq $dot) {
                        $dot = $s.Shapes.AddShape(9, $dxp, $dy, $dw, $dw)
                        $dot.Name = "!!sp_dot_$pi"
                    } else {
                        $dot.Left = $dxp; $dot.Top = $dy; $dot.Width = $dw; $dot.Height = $dw
                    }
                    $dot.Fill.Solid()
                    $dot.Fill.ForeColor.RGB = if ($isAct) { $cAcc1 } else { $cMuted }
                    $dot.Line.Visible = 0
                    $dot.ZOrder(0)
                }
                Set-Morph $s ([single]2.0)
            }

            "macbook_studio_3d" {
                $keep = @("!!PhoneShadow", "!!Phone3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Bg $s "$AssetsDir\bg_alt.png" | Out-Null
                $mSide = if ($sl.side) { [string]$sl.side } else { "right" }

                if ($mSide -eq "left") {
                    $mColX = [single]548; $mColW = [single]368
                    $lapShX = [single]8; $lapX = [single]28; $lapRy = [single]22; $lapRz = [single]2
                } else {
                    $mColX = [single]44; $mColW = [single]366
                    $lapShX = [single]408; $lapX = [single]432; $lapRy = [single]-22; $lapRz = [single]-2
                }

                Add-Text $s "m_ovr" $sl.overline $mColX 40 $mColW 20 $MonoFont 10 $true $cAcc2 1 3 | Out-Null
                Add-Text $s "m_h1" $sl.headline $mColX 64 $mColW 78 $DisplayFont 28 $true $cWhite 1 1 | Out-Null

                $spYs = @(152, 264, 376); $spCols = @($cAcc1, $cAcc2, $cAcc3); $pillars = @($sl.pillars)
                for ($i = 0; $i -lt [Math]::Min(3, $pillars.Count); $i++) {
                    $sy = $spYs[$i]; $cIdx = $i + 1
                    Add-Pic $s "m_card_$cIdx" "$AssetsDir\card_spec_$cIdx.png" $mColX $sy $mColW 96 | Out-Null
                    Add-Text $s "m_ttl_$cIdx" $pillars[$i].title ($mColX + 20) ($sy + 14) ($mColW - 32) 22 $DisplayFont 12.5 $true $spCols[$i] 1 3 | Out-Null
                    Add-Text $s "m_dsc_$cIdx" $pillars[$i].desc ($mColX + 20) ($sy + 40) ($mColW - 32) 48 $BodyFont 11 $false $cLightText 1 1 | Out-Null
                }

                Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" -280 540 200 44 | Out-Null
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb -240 290 148 296 14 -55 -10 | Out-Null

                $shL = Add-Pic $s "!!LaptopShadow" "$AssetsDir\shadow_laptop.png" $lapShX 398 536 84; $shL.ZOrder(0)
                Ensure-Model3D $s "!!Laptop3D" $macbookGlb $lapX 86 498 348 9 $lapRy $lapRz | Out-Null
                Set-Morph $s ([single]2.1)
            }

            "problem_solution_split" {
                $keep = @("!!PhoneShadow", "!!Phone3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Bg $s "$AssetsDir\bg_main.png" | Out-Null
                Add-Text $s "ps_ovr" $sl.overline 42 24 620 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "ps_h1" $sl.headline 42 46 720 44 $DisplayFont 27 $true $cWhite 1 3 | Out-Null

                Add-Pic $s "ps_lbg" "$AssetsDir\card_split_left.png" 38 102 302 404 | Out-Null
                Add-Text $s "ps_ltag" $sl.left.tag 60 120 258 18 $MonoFont 9 $true $cMuted 1 3 | Out-Null
                Add-Text $s "ps_ltl" $sl.left.title 60 140 258 32 $DisplayFont 17 $true $cWhite 1 3 | Out-Null

                $lItems = @($sl.left.items); $rowYs = @(186, 286, 386)
                for ($li = 0; $li -lt [Math]::Min(3, $lItems.Count); $li++) {
                    $ry = $rowYs[$li]; $it = $lItems[$li]
                    Add-Rect $s "ps_lb_$li" 58 ($ry + 8) 3 64 $cMuted 0 -1 0 0 | Out-Null
                    Add-Text $s "ps_lit_$li" $it.title 72 ($ry + 6) 250 24 $SemiFont 12.5 $true $cWhite 1 3 | Out-Null
                    Add-Text $s "ps_lid_$li" $it.desc 72 ($ry + 32) 250 46 $BodyFont 10.5 $false $cMuted 1 1 | Out-Null
                }

                $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 346 420 268 54; $shP.ZOrder(0)
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb 354 114 252 318 12 18 -3 | Out-Null
                $bridge = if ($sl.bridge_label) { [string]$sl.bridge_label } else { "SHIFT" }
                $bPill = Add-CapsulePill $s "ps_bridge" $bridge 432 456 96 32 $cAcc1 0 $cWhite 10 $MonoFont; $bPill.ZOrder(0)

                Add-Pic $s "ps_rbg" "$AssetsDir\card_split_right.png" 620 102 302 404 | Out-Null
                Add-Text $s "ps_rtag" $sl.right.tag 642 120 258 18 $MonoFont 9 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "ps_rtl" $sl.right.title 642 140 258 32 $DisplayFont 17 $true $cWhite 1 3 | Out-Null

                $rItems = @($sl.right.items)
                for ($ri = 0; $ri -lt [Math]::Min(3, $rItems.Count); $ri++) {
                    $ry = $rowYs[$ri]; $it = $rItems[$ri]
                    Add-Rect $s "ps_rb_$ri" 640 ($ry + 8) 4 64 $cAcc1 0 -1 0 0 | Out-Null
                    Add-Text $s "ps_rit_$ri" $it.title 654 ($ry + 6) 250 24 $DisplayFont 13 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "ps_rid_$ri" $it.desc 654 ($ry + 32) 250 46 $BodyFont 10.5 $false $cLightText 1 1 | Out-Null
                }
                Set-Morph $s ([single]1.9)
            }

            "comparison_matrix" {
                $keep = @("!!PhoneShadow", "!!Phone3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Bg $s "$AssetsDir\bg_alt.png" | Out-Null
                Add-Text $s "cm_ovr" $sl.overline 42 24 640 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "cm_h1" $sl.headline 42 46 700 44 $DisplayFont 27 $true $cWhite 1 3 | Out-Null

                $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 28 438 228 50; $shP.ZOrder(0)
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb 36 122 212 326 12 24 -3 | Out-Null

                $cols = @($sl.columns)
                $col1Lbl = if ($cols.Count -gt 0) { [string]$cols[0] } else { "Criterion" }
                $col2Lbl = if ($cols.Count -gt 1) { [string]$cols[1] } else { "Legacy" }
                $col3Lbl = if ($cols.Count -gt 2) { [string]$cols[2] } else { "Alternative" }
                $col4Lbl = if ($cols.Count -gt 3) { [string]$cols[3] } else { $spec.brand.name }

                Add-Rect $s "cm_hc0" 266 106 206 40 $cSurfAlt ([single]0.15) $cBorder ([single]1.0) $cardAdj | Out-Null
                Add-Text $s "cm_ht0" $col1Lbl 278 106 184 40 $MonoFont 9.5 $true $cMuted 1 3 | Out-Null
                Add-Rect $s "cm_hc1" 480 106 134 40 $cSurfAlt ([single]0.15) $cBorder ([single]1.0) $cardAdj | Out-Null
                Add-Text $s "cm_ht1" $col2Lbl 486 106 122 40 $DisplayFont 11 $true $cMuted 2 3 | Out-Null
                Add-Rect $s "cm_hc2" 622 106 134 40 $cSurfAlt ([single]0.15) $cBorder ([single]1.0) $cardAdj | Out-Null
                Add-Text $s "cm_ht2" $col3Lbl 628 106 122 40 $DisplayFont 11 $true $cMuted 2 3 | Out-Null
                Add-Rect $s "cm_hc3" 764 106 154 40 $cAcc1 0 $cAcc1 ([single]1.2) $cardAdj | Out-Null
                Add-Text $s "cm_ht3" $col4Lbl 770 106 142 40 $DisplayFont 12 $true $cWhite 2 3 | Out-Null

                $rows = @($sl.rows)
                for ($ri = 0; $ri -lt [Math]::Min(5, $rows.Count); $ri++) {
                    $ry = [single](156 + $ri * 66); $r = $rows[$ri]
                    Add-Pic $s "cm_rbg_$ri" "$AssetsDir\card_step.png" 266 $ry 490 58 | Out-Null
                    Add-Text $s "cm_rl_$ri" $r.criterion 284 $ry 186 58 $SemiFont 11.5 $true $cWhite 1 3 | Out-Null
                    Add-Text $s "cm_rv1_$ri" $r.col1 486 $ry 122 58 $BodyFont 10.5 $false $cMuted 2 3 | Out-Null
                    Add-Text $s "cm_rv2_$ri" $r.col2 628 $ry 122 58 $BodyFont 10.5 $false $cMuted 2 3 | Out-Null
                    Add-Rect $s "cm_rhi_$ri" 764 $ry 154 58 $cSurface ([single]0.12) $cAcc1 ([single]1.6) $cardAdj | Out-Null
                    Add-Text $s "cm_rv3_$ri" $r.featured 772 $ry 138 58 $DisplayFont 12 $true $cAcc1 2 3 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "chapter_divider_morph" {
                $keep = @("!!PhoneShadow", "!!Phone3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Bg $s "$AssetsDir\bg_finale.png" | Out-Null

                Add-Text $s "cd_ovr" $sl.overline 58 34 520 20 $MonoFont 10 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "cd_num" $sl.chapter_num 54 66 320 140 $HeadingFont 104 $true $cAcc1 1 1 | Out-Null
                Add-Rect $s "cd_bar" 58 226 124 5 $cAcc1 0 -1 0 0 | Out-Null
                Add-Text $s "cd_h1" $sl.headline 58 248 520 112 $DisplayFont 34 $true $cWhite 1 1 | Out-Null
                Add-Text $s "cd_sub" $sl.subtitle 58 368 510 56 $BodyFont 13.5 $false $cLightText 1 1 | Out-Null

                $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 596 418 320 60; $shP.ZOrder(0)
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb 612 64 290 364 16 -32 4 | Out-Null

                Add-Rect $s "cd_line" 58 448 844 ([single]1.0) $cBorder 0 -1 0 0 | Out-Null
                $toc = @($sl.toc)
                for ($ti = 0; $ti -lt [Math]::Min(4, $toc.Count); $ti++) {
                    $tx = [single](58 + $ti * 214)
                    $isCur = [bool]$toc[$ti].active
                    $tCol = if ($isCur) { $cAcc1 } else { $cLightText }
                    Add-Text $s "cd_toc_$ti" $toc[$ti].label $tx 464 200 28 $MonoFont 10 $isCur $tCol 1 3 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "full_bleed_visual" {
                $keep = @("!!PhoneShadow", "!!Phone3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Bg $s "$AssetsDir\s$idx`_visual.png" | Out-Null
                $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 610 384 280 54; $shP.ZOrder(0)
                Ensure-Model3D $s "!!Phone3D" $iphoneGlb 624 54 256 340 12 -24 2 | Out-Null

                if ($sl.overline) { Add-CapsulePill $s "fb_ovr" $sl.overline 46 32 320 28 $cObsidian ([single]0.15) $cAcc1 9.5 $MonoFont | Out-Null }
                Add-Pic $s "fb_bar" "$AssetsDir\card_bento_strip.png" 46 406 868 102 | Out-Null
                Add-Text $s "fb_h1" $sl.headline 74 418 590 34 $DisplayFont 21 $true $cWhite 1 3 | Out-Null
                Add-Text $s "fb_cap" $sl.caption 74 454 590 38 $BodyFont 12 $false $cLightText 1 1 | Out-Null
                if ($sl.badge) { Add-CapsulePill $s "fb_badge" $sl.badge 696 440 194 34 $cAcc1 0 $cWhite 10 $MonoFont | Out-Null }
                Set-Morph $s ([single]1.8)
            }

            "closing_finale" {
                $style = if ($sl.style) { [string]$sl.style } else { "editorial_statement" }
                if ($style -eq "dual_3d") {
                    $keep = @("!!LaptopShadow", "!!Laptop3D", "!!PhoneShadow", "!!Phone3D")
                    $s = Start-MorphSlide $pres $prevSlide $keep
                    Add-Bg $s "$AssetsDir\bg_finale.png" | Out-Null
                    Add-Pic $s "f_logo" "$AssetsDir\logo_badge.png" 48 74 62 62 | Out-Null
                    Add-Text $s "f_title" $spec.brand.name 48 148 312 54 $DisplayFont 40 $true $cWhite 1 1 | Out-Null
                    Add-Text $s "f_tag" $sl.tagline 48 206 310 28 $DisplayFont 15.5 $true $cAcc1 1 3 | Out-Null
                    Add-Rect $s "f_div" 48 246 56 ([single]2.5) $cAcc1 0 -1 0 0 | Out-Null
                    Add-Text $s "f_sum" $sl.summary 48 264 310 112 $BodyFont 12.5 $false $cLightText 1 1 | Out-Null
                    Add-CapsulePill $s "f_cta" $sl.cta 48 396 294 38 $cAcc1 0 $cWhite 11 $DisplayFont | Out-Null

                    $shL = Add-Pic $s "!!LaptopShadow" "$AssetsDir\shadow_laptop.png" 468 396 480 78; $shL.ZOrder(0)
                    Ensure-Model3D $s "!!Laptop3D" $macbookGlb 486 114 452 316 6 -14 0 | Out-Null
                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 356 442 200 46; $shP.ZOrder(0)
                    $fW = if ($isPhoneHero) { [single]144 } else { [single]210 }
                    $fH = if ($isPhoneHero) { [single]288 } else { [single]240 }
                    $fY = if ($isPhoneHero) { [single]176 } else { [single]214 }
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 364 $fY $fW $fH 6 20 -3 | Out-Null
                }
                elseif ($style -eq "executive_cta") {
                    $keep = @("!!PhoneShadow", "!!Phone3D")
                    $s = Start-MorphSlide $pres $prevSlide $keep
                    Add-Bg $s "$AssetsDir\bg_finale.png" | Out-Null

                    Add-Text $s "fe_ovr" $sl.overline 52 34 560 20 $MonoFont 10 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "fe_h1" $sl.headline 52 58 560 72 $DisplayFont 34 $true $cWhite 1 3 | Out-Null
                    Add-Text $s "fe_sum" $sl.summary 52 134 540 44 $BodyFont 13 $false $cLightText 1 1 | Out-Null

                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_phone.png" 648 202 256 44; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 664 24 226 196 14 -24 3 | Out-Null

                    $actions = @($sl.actions); $acCols = @($cAcc1, $cAcc2, $cAcc3)
                    for ($ai = 0; $ai -lt [Math]::Min(3, $actions.Count); $ai++) {
                        $ax = [single](52 + $ai * 292); $cIdx = $ai + 1
                        Add-Pic $s "fe_bg_$ai" "$AssetsDir\card_bento_tile_$cIdx.png" $ax 224 272 178 | Out-Null
                        Add-Text $s "fe_num_$ai" $actions[$ai].num ($ax + 22) 244 80 24 $MonoFont 13 $true $acCols[$ai] 1 3 | Out-Null
                        Add-Text $s "fe_ttl_$ai" $actions[$ai].title ($ax + 22) 272 228 28 $DisplayFont 15 $true $cWhite 1 3 | Out-Null
                        Add-Text $s "fe_dsc_$ai" $actions[$ai].desc ($ax + 22) 306 228 76 $BodyFont 11.5 $false $cLightText 1 1 | Out-Null
                    }

                    Add-Pic $s "fe_fbar" "$AssetsDir\card_step.png" 52 428 856 64 | Out-Null
                    Add-Text $s "fe_fbrand" "$($spec.brand.name)  |  $($sl.tagline)" 82 444 480 32 $SemiFont 13 $true $cWhite 1 3 | Out-Null
                    Add-CapsulePill $s "fe_cta" $sl.cta 644 442 240 36 $cAcc1 0 $cWhite 10.5 $MonoFont | Out-Null
                }
                else {
                    $keep = @("!!PhoneShadow", "!!Phone3D")
                    $s = Start-MorphSlide $pres $prevSlide $keep
                    Add-Bg $s "$AssetsDir\bg_finale.png" | Out-Null
                    Add-Pic $s "fs_logo" "$AssetsDir\logo_badge.png" 52 48 62 62 | Out-Null
                    Add-Text $s "fs_ovr" $sl.overline 128 66 440 22 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "fs_h1" $sl.headline 52 128 500 112 $DisplayFont 38 $true $cWhite 1 1 | Out-Null
                    Add-Rect $s "fs_rule" 52 252 120 ([single]2.5) $cAcc1 0 -1 0 0 | Out-Null
                    Add-Text $s "fs_sub" $sl.summary 52 270 490 66 $BodyFont 13.5 $false $cLightText 1 1 | Out-Null

                    $metrics = @($sl.metrics); $mCols = @($cAcc1, $cAcc2, $cAcc3)
                    for ($mi = 0; $mi -lt [Math]::Min(3, $metrics.Count); $mi++) {
                        $mx = [single](52 + $mi * 166)
                        Add-DarkKpiBox $s "fs_mk_$mi" $metrics[$mi].value $metrics[$mi].label $mx 352 152 72 $cObsidian $mCols[$mi] $cLightText | Out-Null
                    }
                    Add-CapsulePill $s "fs_cta" $sl.cta 52 446 286 40 $cAcc1 0 $cWhite 10.5 $DisplayFont | Out-Null

                    $shP = Add-Pic $s "!!PhoneShadow" "$AssetsDir\shadow_laptop.png" 546 432 376 68; $shP.ZOrder(0)
                    Ensure-Model3D $s "!!Phone3D" $iphoneGlb 568 62 336 382 12 -24 2 | Out-Null
                }
                Set-Morph $s ([single]2.0)
            }

            default {
                throw "Unsupported slide type '$st' at index $idx"
            }
        }

        # Pre-stage staggered off-canvas callouts/tiles when next slide uses macro_dolly_3d or kpi_bento
        if ($nextType -eq "macro_dolly_3d" -and $st -ne "cover_hero") {
            for ($mi = 0; $mi -lt 3; $mi++) {
                Add-Pic $s "!!md_card_$mi" "$AssetsDir\card_spec_$($mi + 1).png" (-460 - $mi * 140) (142 + $mi * 116) 384 96 | Out-Null
            }
        }

        $prevSlide = $s
        $prevType  = $st
    }

    if (Test-Path $OutputPptx) { Remove-Item $OutputPptx -Force -ErrorAction SilentlyContinue }
    $pres.SaveAs($OutputPptx)
    Write-Host ("Saved deck cleanly -> " + $OutputPptx)

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
