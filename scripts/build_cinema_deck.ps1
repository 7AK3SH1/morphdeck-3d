# Copyright 2026 7AK3SH1
# Licensed under the Apache License, Version 2.0 (see LICENSE and NOTICE).
# Master Cinema-Morph v3 PowerPoint COM Builder for morphdeck-3d

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
        $existing.Fill.Solid(); $existing.Fill.ForeColor.RGB = $fillBgr; $existing.Fill.Transparency = $fillTrans
        if ($borderBgr -ge 0 -and $borderWeight -gt 0) {
            $existing.Line.Visible = -1; $existing.Line.ForeColor.RGB = $borderBgr; $existing.Line.Weight = $borderWeight
        } else { $existing.Line.Visible = 0 }
        return $existing
    }
    $stype = if ($cornerAdj -gt 0.0) { 5 } else { 1 }
    $sh = $slide.Shapes.AddShape($stype, $x, $y, $w, $h)
    $sh.Name = $name
    if ($cornerAdj -gt 0.0) { $sh.Adjustments.Item(1) = $cornerAdj }
    $sh.Fill.Solid(); $sh.Fill.ForeColor.RGB = $fillBgr; $sh.Fill.Transparency = $fillTrans
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
        $sh.Visible = -1; $sh.Blur = [single]6; $sh.OffsetX = [single]0; $sh.OffsetY = [single]3; $sh.Transparency = [single]0.48
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

function Add-KpiBox($slide, [string]$name, [string]$val, [string]$lbl, [single]$x, [single]$y, [single]$w, [single]$h, [int]$bgBgr, [int]$accBgr, [int]$txtBgr, [int]$bdrBgr) {
    $sh = Add-Rect $slide $name $x $y $w $h $bgBgr ([single]0.12) $bdrBgr ([single]1.2) ([single]0.14)
    $tf = $sh.TextFrame2; $tf.WordWrap = -1; $tf.AutoSize = 0; $tf.VerticalAnchor = 3
    $tf.MarginLeft = [single]10; $tf.MarginRight = [single]10
    $tr = $tf.TextRange; $tr.Text = "$val`r$lbl"
    $p1 = $tr.Paragraphs(1); $p1.Font.Name = $DisplayFont; $p1.Font.Size = [single]23; $p1.Font.Bold = -1
    $p1.Font.Fill.Solid(); $p1.Font.Fill.ForeColor.RGB = $accBgr; $p1.ParagraphFormat.Alignment = 2
    $p2 = $tr.Paragraphs(2); $p2.Font.Name = $SemiFont; $p2.Font.Size = [single]9.5; $p2.Font.Bold = 0
    $p2.Font.Fill.Solid(); $p2.Font.Fill.ForeColor.RGB = $txtBgr; $p2.ParagraphFormat.Alignment = 2
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

    for ($si = 0; $si -lt $slides.Count; $si++) {
        $sl = $slides[$si]
        $idx = $si + 1
        $st = [string]$sl.type
        $nextSl = if (($si + 1) -lt $slides.Count) { $slides[$si + 1] } else { $null }
        $nextType = if ($null -ne $nextSl) { [string]$nextSl.type } else { "" }
        Write-Host ("Building Slide {0}/{1} ({2} | universe={3})..." -f $idx, $slides.Count, $st, $spec.universe)

        switch ($st) {
            "newspaper_dive" {
                $stage = if ($sl.stage) { [string]$sl.stage } else { "frontpage" }
                $keep = @()
                if ($prevType -eq "newspaper_dive") { $keep += "!!news_sheet" }
                $s = Start-MorphSlide $pres $prevSlide $keep

                if ($stage -eq "collage") {
                    $ns = Add-Pic $s "!!news_sheet" "$AssetsDir\news_sheet.png" -1180 -680 3456 1944; $ns.ZOrder(1)
                    Add-Rect $s "ns_scrim" 0 0 960 540 $cBg ([single]0.38) -1 0 0 | Out-Null
                    Add-Pic $s "!!col_cutout_L" $sl._left_cutout 32 94 336 446 | Out-Null
                    Add-Pic $s "!!col_cutout_R" $sl._right_cutout 468 156 410 384 | Out-Null
                    Add-Pic $s "!!col_prop" $sl._prop_cutout 172 34 196 112 | Out-Null

                    Add-Pic $s "ns_panel" "$AssetsDir\card_panel.png" 496 36 420 202 | Out-Null
                    Add-Text $s "ns_ovr" $sl.overline 520 54 372 18 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "ns_h1" $sl.headline 520 76 372 64 $DisplayFont 22 $true $cAcc1 1 1 | Out-Null
                    Add-Text $s "ns_body" $sl.body 520 144 372 78 $BodyFont 11.5 $false $cInk 1 1 | Out-Null
                    Set-Morph $s ([single]2.0)
                } elseif ($stage -eq "masthead_spread") {
                    $ns = Add-Pic $s "!!news_sheet" "$AssetsDir\news_sheet.png" -680 -320 2400 1350; $ns.ZOrder(1)
                    Add-Pic $s "ns_cloud" "$AssetsDir\ep_paper_cloud.png" 160 156 640 248 | Out-Null
                    Add-Text $s "ns_ovr" $sl.overline 210 184 540 22 $MonoFont 10 $true $cAcc1 2 3 | Out-Null
                    Add-Text $s "ns_h1" $sl.headline 200 212 560 88 $DisplayFont 32 $true $cInk 2 3 | Out-Null
                    Add-Text $s "ns_body" $sl.body 220 306 520 76 $BodyFont 13 $false $cMuted 2 1 | Out-Null
                    Set-Morph $s ([single]1.9)
                } else {
                    $ns = Add-Pic $s "!!news_sheet" "$AssetsDir\news_sheet.png" 0 0 960 540; $ns.ZOrder(1)
                    if ($sl.badge) { Add-CapsulePill $s "ns_badge" $sl.badge 668 486 248 30 $cObsidian ([single]0.10) $cWhite 9 $MonoFont | Out-Null }
                    Set-Morph $s ([single]1.6)
                }
            }

            "editorial_print_overlay" {
                $keep = @("!!news_sheet")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $ns = Add-Pic $s "!!news_sheet" "$AssetsDir\news_sheet.png" -280 -180 1440 810; $ns.ZOrder(1)
                Add-Rect $s "ep_scrim" 0 0 960 540 $cBg ([single]0.22) -1 0 0 | Out-Null

                # Organic wavy ink drips hanging from the top edge
                Add-Pic $s "ep_drips" "$AssetsDir\ep_drips.png" 0 0 960 190 | Out-Null
                Add-Pic $s "!!ep_port" $sl._portrait_cutout 46 66 352 474 | Out-Null

                # Soft translucent feathered paper cloud instead of rigid rectangular card
                Add-Pic $s "ep_cloud" "$AssetsDir\ep_paper_cloud.png" 384 108 552 372 | Out-Null
                Add-Text $s "ep_ovr" $sl.overline 438 148 446 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "ep_h1" $sl.headline 438 176 446 72 $DisplayFont 26 $true $cAcc1 1 1 | Out-Null
                Add-Text $s "ep_body" $sl.body 438 256 446 140 $BodyFont 12.5 $false $cInk 1 1 | Out-Null
                if ($sl.badge) { Add-CapsulePill $s "ep_bdg" $sl.badge 438 408 214 30 $cAcc1 0 $cWhite 9.5 $MonoFont | Out-Null }
                Set-Morph $s ([single]1.9)
            }

            "odometer_counter" {
                $keep = @("!!odo_curtain")
                $targetStr = [string]$sl.target_value
                $chars = $targetStr.ToCharArray()
                if ($prevType -eq "odometer_counter" -or $null -ne (Find-Shape $prevSlide "!!odo_col_0")) {
                    for ($ci = 0; $ci -lt $chars.Count; $ci++) {
                        if ($chars[$ci] -match '\d') { $keep += "!!odo_col_$ci" }
                    }
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
                        $sep = Add-Text $s "odo_sep_$ci" ([string]$ch) $curX $winY 32 88 $HeadingFont 66 $true $cWhite 2 3
                        $sep.ZOrder(0)
                        $curX += 32
                    }
                }

                $mTop = Add-Rect $s "odo_mask_top" -4 -4 968 220 $odoBg 0 -1 0 0; $mTop.ZOrder(0)
                $mBot = Add-Rect $s "odo_mask_bot" -4 304 968 242 $odoBg 0 -1 0 0; $mBot.ZOrder(0)

                $silPath = if ($sl._silhouette_cutout) { [string]$sl._silhouette_cutout } else { "$AssetsDir\odo_silhouette.png" }
                $sil = Add-Pic $s "odo_sil" $silPath 0 378 960 162; $sil.ZOrder(0)
                $ovr = Add-Text $s "odo_ovr" $sl.overline 120 144 720 34 $DisplayFont 22 $true $cAcc1 2 3; $ovr.ZOrder(0)
                $sub = Add-Text $s "odo_sub" $sl.sublabel 120 322 720 34 $DisplayFont 24 $true $cWhite 2 3; $sub.ZOrder(0)
                Set-Morph $s ([single]2.1)
            }

            "flyby_wipe" {
                $keep = @("!!flyby_obj", "!!flyby_cloud")
                if ($heroGlb -and $sl.use_hero_3d) { $keep += @("!!HeroShadow", "!!Hero3D") }
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bgP = if ($sl._bg_image) { [string]$sl._bg_image } else { "$AssetsDir\bg_base.png" }
                $bg = Add-Pic $s "fb_bg" $bgP 0 0 960 540; $bg.ZOrder(1)

                $fc = Add-Pic $s "!!flyby_cloud" "$AssetsDir\flyby_cloud.png" 980 -220 760 440; $fc.ZOrder(0)
                if ($sl._flyby_cutout) {
                    $fo = Add-Pic $s "!!flyby_obj" $sl._flyby_cutout 1160 -320 460 280
                    $fo.Rotation = [single]-14; $fo.ZOrder(0)
                }

                $txtCol = if ($sl._bg_image) { $cWhite } else { $cInk }
                if ($sl._bg_image) {
                    Add-Rect $s "fb_scrim" 38 28 534 256 $cObsidian ([single]0.34) -1 0 ([single]0.08) | Out-Null
                }
                Add-Text $s "fb_ovr" $sl.overline 56 44 500 20 $MonoFont 10 $true $cAcc1 1 3 | Out-Null
                $fh1 = Add-Text $s "fb_h1" $sl.headline 56 72 500 88 $DisplayFont 32 $true $txtCol 1 1; Add-TextShadow $fh1
                Add-Text $s "fb_body" $sl.body 56 172 480 104 $BodyFont 13 $false $txtCol 1 1 | Out-Null

                if ($heroGlb -and ($sl.use_hero_3d -ne $false)) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 552 414 352 60; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 560 72 340 360 12 -24 2 | Out-Null
                } elseif ($sl._hero_cutout) {
                    Add-Pic $s "!!HeroCut" $sl._hero_cutout 548 56 368 416 | Out-Null
                }
                Set-Morph $s ([single]1.7)
            }

            "cutout_map_stage" {
                $keep = @()
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bgP = if ($sl._bg_image) { [string]$sl._bg_image } else { "$AssetsDir\bg_base.png" }
                $bg = Add-Pic $s "cm_bg" $bgP 0 0 960 540; $bg.ZOrder(1)

                if ($sl._map_image) {
                    Add-Pic $s "cm_map" $sl._map_image 36 52 536 446 | Out-Null
                }
                Add-Pic $s "cm_panel" "$AssetsDir\card_panel.png" 588 64 334 412 | Out-Null
                Add-Text $s "cm_ovr" $sl.overline 614 86 282 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "cm_h1" $sl.headline 614 114 282 78 $DisplayFont 28 $true $cAcc1 1 1 | Out-Null
                Add-Text $s "cm_body" $sl.body 614 204 282 156 $BodyFont 12 $false $cInk 1 1 | Out-Null
                $kpis = @($sl.kpis)
                for ($ki = 0; $ki -lt [Math]::Min(2, $kpis.Count); $ki++) {
                    Add-KpiBox $s "cm_k_$ki" $kpis[$ki].value $kpis[$ki].label (614 + $ki * 144) 374 134 74 $cSurface $cAcc1 $cInk $cBorder | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "poster_depth_reveal" {
                $keep = @("!!poster_band", "!!poster_word", "!!poster_hero")
                $use3d = ($null -eq $sl._hero_cutout -and $heroGlb -and ($sl.use_hero_3d -ne $false))
                if ($use3d) { $keep += @("!!HeroShadow", "!!Hero3D") }
                $s = Start-MorphSlide $pres $prevSlide $keep

                # Stage museum room far zoomed in BEFORE !!frame_art if transitioning into picture_frame_pullback
                if ($nextType -eq "picture_frame_pullback") {
                    $wallNext = if ($nextSl._wall_image) { [string]$nextSl._wall_image } else { "$AssetsDir\bg_alt.png" }
                    $wr = Add-Pic $s "!!wall_room" $wallNext -1632 -918 4224 2376; $wr.ZOrder(1)
                }

                $bgP = if ($sl._poster_bg_image) { [string]$sl._poster_bg_image } elseif ($sl._bg_image) { [string]$sl._bg_image } else { "$AssetsDir\bg_alt.png" }
                $pbg = Add-Pic $s "!!frame_art" $bgP 0 0 960 540; $pbg.ZOrder(0)

                $band = Add-Rect $s "!!poster_band" 348 0 264 540 $cAcc1 0 -1 0 0; $band.ZOrder(0)
                $gw = if ($sl.giant_word) { [string]$sl.giant_word } else { "2026" }
                $pw = Add-Text $s "!!poster_word" $gw 60 52 840 168 $HeadingFont 124 $true $cWhite 2 3; Add-TextShadow $pw; $pw.ZOrder(0)

                if ($sl._hero_cutout) {
                    $ph = Add-Pic $s "!!poster_hero" $sl._hero_cutout 170 76 620 464; $ph.ZOrder(0)
                } elseif ($use3d) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 290 426 380 70; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 288 76 384 374 12 -22 2 | Out-Null
                }

                if ($sl._fg_left_cutout)  { $fl = Add-Pic $s "poster_fg_L" $sl._fg_left_cutout -20 290 250 250; $fl.ZOrder(0) }
                if ($sl._fg_right_cutout) { $fr = Add-Pic $s "poster_fg_R" $sl._fg_right_cutout 730 290 250 250; $fr.ZOrder(0) }

                if ($nextType -eq "picture_frame_pullback" -and $nextSl._hollow_frame) {
                    # Stage hollow museum frame OUTSIDE 960x540 viewport so poster layers remain 100% unobstructed
                    $hf = Add-Pic $s "!!hollow_frame" $nextSl._hollow_frame -240 -180 1440 900; $hf.ZOrder(0)
                }
                Set-Morph $s ([single]2.0)
            }

            "picture_frame_pullback" {
                $keep = @("!!wall_room", "!!frame_art", "!!hollow_frame", "!!poster_band", "!!poster_word", "!!poster_hero")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $wallP = if ($sl._wall_image) { [string]$sl._wall_image } else { "$AssetsDir\bg_alt.png" }
                $wbg = Add-Pic $s "!!wall_room" $wallP 0 0 960 540; $wbg.ZOrder(1)

                $fx = if ($null -ne $sl.frame_x) { [single]$sl.frame_x } else { [single]348 }
                $fy = if ($null -ne $sl.frame_y) { [single]$sl.frame_y } else { [single]42 }
                $fw = if ($null -ne $sl.frame_w) { [single]$sl.frame_w } else { [single]264 }
                $fh = if ($null -ne $sl.frame_h) { [single]$sl.frame_h } else { [single]148 }

                $artSrc = if ($sl._framed_art) { [string]$sl._framed_art } elseif ($null -ne (Find-Shape $s "!!frame_art")) { "" } else { "$AssetsDir\bg_alt.png" }
                if ($artSrc) {
                    $fa = Add-Pic $s "!!frame_art" $artSrc $fx $fy $fw $fh; $fa.ZOrder(0)
                } elseif ($null -ne (Find-Shape $s "!!frame_art")) {
                    $fa = Find-Shape $s "!!frame_art"
                    $fa.Left = $fx; $fa.Top = $fy; $fa.Width = $fw; $fa.Height = $fh; $fa.ZOrder(0)
                }

                # Shrink carried poster_band, giant word & miniature subject cleanly inside the framed canvas
                $pb = Find-Shape $s "!!poster_band"
                if ($null -ne $pb) {
                    $pb.Left = [single]($fx + $fw * 0.362); $pb.Top = $fy
                    $pb.Width = [single]($fw * 0.276); $pb.Height = $fh; $pb.ZOrder(0)
                }
                $pw = Find-Shape $s "!!poster_word"
                if ($null -ne $pw) {
                    $pw.Left = [single]($fx + $fw * 0.06); $pw.Top = [single]($fy + $fh * 0.08)
                    $pw.Width = [single]($fw * 0.88); $pw.Height = [single]($fh * 0.34)
                    $pw.TextFrame2.TextRange.Font.Size = [single]32
                    $pw.ZOrder(0)
                }
                $ph = Find-Shape $s "!!poster_hero"
                if ($null -ne $ph) {
                    $ph.Left = [single]($fx + $fw * 0.22); $ph.Top = [single]($fy + $fh * 0.14)
                    $ph.Width = [single]($fw * 0.56); $ph.Height = [single]($fh * 0.82); $ph.ZOrder(0)
                } elseif (Test-Path "$AssetsDir\hero_cutout_auto.png") {
                    $autoMini = Add-Pic $s "pf_mini_hero" "$AssetsDir\hero_cutout_auto.png" ([single]($fx + $fw * 0.26)) ([single]($fy + $fh * 0.18)) ([single]($fw * 0.48)) ([single]($fh * 0.78))
                    if ($autoMini) { $autoMini.ZOrder(0) }
                }

                if ($sl._hollow_frame) {
                    # Exact hollow aperture wrap around ($fx, $fy, $fw, $fh)
                    $hfX = [single]($fx - ($fw * 0.0755))
                    $hfY = [single]($fy - ($fh * 0.152))
                    $hfW = [single]($fw / 0.86875)
                    $hfH = [single]($fh / 0.76667)
                    $hf = Add-Pic $s "!!hollow_frame" $sl._hollow_frame $hfX $hfY $hfW $hfH; $hf.ZOrder(0)
                }

                if ($sl._room_subject_cutout) {
                    $rs = Add-Pic $s "pf_room_subj" $sl._room_subject_cutout 470 156 468 384; $rs.ZOrder(0)
                }

                Add-Rect $s "pf_scrim" 44 232 396 260 $cObsidian ([single]0.34) -1 0 ([single]0.10) | Out-Null
                Add-Text $s "pf_ovr" $sl.overline 62 246 360 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                $pfh = Add-Text $s "pf_h1" $sl.headline 62 272 360 92 $DisplayFont 30 $true $cWhite 1 1; Add-TextShadow $pfh
                Add-Text $s "pf_sub" $sl.subtitle 62 370 356 54 $BodyFont 12.5 $false $cWhite 1 1 | Out-Null
                if ($sl.cta) { Add-CapsulePill $s "pf_cta" $sl.cta 62 436 256 36 $cAcc1 0 $cWhite 10 $DisplayFont | Out-Null }
                Set-Morph $s ([single]2.2)
            }

            "real_portal_flythrough" {
                $stage = if ($sl.stage) { [string]$sl.stage } else { "exterior" }
                $keep = @("!!portal_bg", "!!portal_inner", "!!portal_outer", "!!HeroShadow", "!!Hero3D", "!!pt_card_next")
                $s = Start-MorphSlide $pres $prevSlide $keep

                if ($stage -eq "interior") {
                    $pbg = Add-Pic $s "!!portal_bg" "$AssetsDir\portal_bg.png" -64 -36 1088 612; $pbg.ZOrder(1)
                    $pout = Add-Pic $s "!!portal_outer" "$AssetsDir\portal_outer.png" -1824 -1026 4608 2592; $pout.ZOrder(0)
                    $pin = Add-Pic $s "!!portal_inner" "$AssetsDir\portal_inner.png" 0 0 960 540; $pin.ZOrder(0)

                    $pcard = Add-Pic $s "!!pt_card_next" "$AssetsDir\card_panel.png" 56 102 426 336; $pcard.ZOrder(0)
                    Add-Text $s "pt_ovr" $sl.overline 80 122 378 18 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "pt_h1" $sl.headline 80 146 378 68 $DisplayFont 24 $true $cInk 1 1 | Out-Null
                    Add-Text $s "pt_body" $sl.body 80 220 378 96 $BodyFont 11.5 $false $cMuted 1 1 | Out-Null
                    $kpis = @($sl.kpis)
                    for ($ki = 0; $ki -lt [Math]::Min(2, $kpis.Count); $ki++) {
                        Add-KpiBox $s "pt_k_$ki" $kpis[$ki].value $kpis[$ki].label (80 + $ki * 192) 334 180 76 $cSurface $cAcc1 $cInk $cBorder | Out-Null
                    }
                    if ($heroGlb) {
                        $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 524 416 340 60; $sh.ZOrder(0)
                        Ensure-Model3D $s "!!Hero3D" $heroGlb 542 78 308 348 12 -22 2 | Out-Null
                    } elseif ($sl._hero_cutout) {
                        Add-Pic $s "!!HeroCut" $sl._hero_cutout 542 78 308 348 | Out-Null
                    }
                    Set-Morph $s ([single]2.1)
                } elseif ($stage -eq "horizon") {
                    $pbg = Add-Pic $s "!!portal_bg" "$AssetsDir\portal_bg.png" -120 -68 1200 675; $pbg.ZOrder(1)
                    $pout = Add-Pic $s "!!portal_outer" "$AssetsDir\portal_outer.png" -1824 -1026 4608 2592; $pout.ZOrder(0)
                    $pin = Add-Pic $s "!!portal_inner" "$AssetsDir\portal_inner.png" -1632 -918 4224 2376; $pin.ZOrder(0)
                    Add-Text $s "pt_ovr" $sl.overline 54 48 460 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    $ph1 = Add-Text $s "pt_h1" $sl.headline 54 74 460 96 $DisplayFont 32 $true $cWhite 1 1; Add-TextShadow $ph1
                    Add-Text $s "pt_body" $sl.body 54 180 440 100 $BodyFont 12.5 $false $cWhite 1 1 | Out-Null
                    $kpis = @($sl.kpis)
                    for ($ki = 0; $ki -lt [Math]::Min(3, $kpis.Count); $ki++) {
                        Add-KpiBox $s "pt_k_$ki" $kpis[$ki].value $kpis[$ki].label (54 + $ki * 158) 396 146 78 $cObsidian $cAcc1 $cWhite $cBorder | Out-Null
                    }
                    if ($heroGlb) {
                        $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 524 424 372 64; $sh.ZOrder(0)
                        Ensure-Model3D $s "!!Hero3D" $heroGlb 542 64 342 372 14 24 -3 | Out-Null
                    }
                    Set-Morph $s ([single]2.0)
                } else {
                    $pbg = Add-Pic $s "!!portal_bg" "$AssetsDir\portal_bg.png" 0 0 960 540; $pbg.ZOrder(1)
                    if ($heroGlb) {
                        Ensure-Model3D $s "!!Hero3D" $heroGlb 392 152 176 210 16 -28 -4 | Out-Null
                    }
                    $pin = Add-Pic $s "!!portal_inner" "$AssetsDir\portal_inner.png" 272 102 416 336; $pin.ZOrder(0)
                    $pout = Add-Pic $s "!!portal_outer" "$AssetsDir\portal_outer.png" 0 0 960 540; $pout.ZOrder(0)
                    Add-Rect $s "pt_top_scrim" 160 10 640 76 $cObsidian ([single]0.38) -1 0 ([single]0.18) | Out-Null
                    Add-Text $s "pt_ovr" $sl.overline 180 18 600 18 $MonoFont 9.5 $true $cAcc1 2 3 | Out-Null
                    $ph1 = Add-Text $s "pt_h1" $sl.headline 180 38 600 42 $DisplayFont 24 $true $cWhite 2 3; Add-TextShadow $ph1
                    if ($sl.bottom_cue) {
                        $cp = Add-CapsulePill $s "pt_cue" $sl.bottom_cue 290 482 380 34 $cObsidian ([single]0.14) $cWhite 10 $MonoFont
                        $cp.Line.Visible = -1; $cp.Line.ForeColor.RGB = $cAcc1; $cp.Line.Weight = [single]1.2
                    }
                    if ($nextType -eq "real_portal_flythrough") {
                        Add-Pic $s "!!pt_card_next" "$AssetsDir\card_panel.png" -520 102 426 336 | Out-Null
                    }
                    Set-Morph $s ([single]2.0)
                }
            }

            "refracted_spotlight_dock" {
                $dockIdx = if ($null -ne $sl.dock_index) { [int]$sl.dock_index } else { 0 }
                $dockXs = @(66, 282, 498, 714); $gliderXs = @(54, 270, 486, 702)
                $keep = @("!!HeroShadow", "!!Hero3D")
                if ($prevType -eq "refracted_spotlight_dock") {
                    $prevIdx = $idx - 1
                    $keep += @("!!sp_bg_$idx", "!!sp_dot_0", "!!sp_dot_1", "!!sp_tag_$prevIdx", "!!sp_title_$prevIdx", "!!sp_sub_$prevIdx", "!!sp_body_$prevIdx", "!!DockGlider", "!!DockThumb1", "!!DockThumb2", "!!DockThumb3", "!!DockThumb4", "!!DockLbl1", "!!DockLbl2", "!!DockLbl3", "!!DockLbl4", "!!DockSub1", "!!DockSub2", "!!DockSub3", "!!DockSub4")
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
                    Add-Text $s "!!sp_tag_$nextIdx" $nxSp.tag -460 48 540 20 $MonoFont 10 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "!!sp_title_$nextIdx" $nxSp.title -560 74 550 64 $DisplayFont 42 $true $cWhite 1 1 | Out-Null
                    Add-Text $s "!!sp_sub_$nextIdx" $nxSp.subtitle -680 142 560 26 $SemiFont 13 $false $cAcc3 1 3 | Out-Null
                    Add-Text $s "!!sp_body_$nextIdx" $nxSp.body -800 180 520 112 $BodyFont 13.5 $false $cWhite 1 1 | Out-Null
                }

                $sp = $sl.spotlight
                Add-Rect $s "sp_scrim" 32 34 566 266 $cObsidian ([single]0.36) -1 0 ([single]0.08) | Out-Null
                Add-Text $s "!!sp_tag_$idx" $sp.tag 50 46 530 20 $MonoFont 10 $true $cAcc1 1 3 | Out-Null
                $spt = Add-Text $s "!!sp_title_$idx" $sp.title 50 70 530 78 $DisplayFont 32 $true $cWhite 1 1; Add-TextShadow $spt
                Add-Text $s "!!sp_sub_$idx" $sp.subtitle 50 156 530 26 $SemiFont 13 $false $cAcc3 1 3 | Out-Null
                Add-Text $s "!!sp_body_$idx" $sp.body 50 188 520 100 $BodyFont 13 $false $cWhite 1 1 | Out-Null

                if ($heroGlb) {
                    $rySp = if ($dockIdx % 2 -eq 0) { [single]-16 } else { [single]22 }
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 644 304 268 46; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 652 38 252 276 10 $rySp 2 | Out-Null
                } elseif ($sl._portrait_cutout) {
                    Add-Pic $s "sp_port" $sl._portrait_cutout 646 36 254 292 | Out-Null
                }

                $refractDock = "$AssetsDir\s$idx`_dock_refract.png"
                $dockPath = if (Test-Path $refractDock) { $refractDock } else { "$AssetsDir\dock_bg.png" }
                $gd = Find-Shape $s "!!GlassDock"; if ($null -ne $gd) { $gd.Delete() }
                Add-Pic $s "!!GlassDock" $dockPath 40 386 880 128 | Out-Null
                $gld = Add-Pic $s "!!DockGlider" "$AssetsDir\dock_glider.png" $gliderXs[[Math]::Min(3, [Math]::Max(0, $dockIdx))] 396 208 108; $gld.ZOrder(0)

                $dockItems = @($sl.dock)
                for ($di = 0; $di -lt 4; $di++) {
                    $it = if ($di -lt $dockItems.Count) { $dockItems[$di] } else { @{ title = "Item $($di+1)"; subtitle = "" } }
                    $dn = $di + 1; $dx = $dockXs[$di]
                    $th = Add-Pic $s "!!DockThumb$dn" "$AssetsDir\thumb_$dn.png" $dx 416 76 60; if ($th) { $th.ZOrder(0) }
                    $tl = Add-Text $s "!!DockLbl$dn" $it.title ($dx + 84) 418 110 24 $DisplayFont 13 $true $cWhite 1 3; $tl.ZOrder(0)
                    $ts = Add-Text $s "!!DockSub$dn" $it.subtitle ($dx + 84) 444 110 22 $BodyFont 11 $false $cMuted 1 1; $ts.ZOrder(0)
                }
                for ($pi = 0; $pi -lt 2; $pi++) {
                    $isAct = ($pi -eq ($dockIdx % 2)); $dw = if ($isAct) { [single]14 } else { [single]7 }; $dxp = if ($isAct) { [single]930 } else { [single]933.5 }
                    $dot = Find-Shape $s "!!sp_dot_$pi"
                    if ($null -eq $dot) { $dot = $s.Shapes.AddShape(9, $dxp, [single](222 + $pi * 24), $dw, $dw); $dot.Name = "!!sp_dot_$pi" }
                    else { $dot.Left = $dxp; $dot.Top = [single](222 + $pi * 24); $dot.Width = $dw; $dot.Height = $dw }
                    $dot.Fill.Solid(); $dot.Fill.ForeColor.RGB = if ($isAct) { $cAcc1 } else { $cMuted }; $dot.Line.Visible = 0; $dot.ZOrder(0)
                }
                Set-Morph $s ([single]2.0)
            }

            "parchment_3d_split" {
                $keep = @("!!HeroShadow", "!!Hero3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bgP = "$AssetsDir\bg_base.png"
                $bg = Add-Pic $s "bg_slide" $bgP 0 0 960 540; $bg.ZOrder(1)
                $lowPath = "$AssetsDir\s$idx`_split_lower.png"
                if (Test-Path $lowPath) { $lw = Add-Pic $s "!!SplitLower" $lowPath 0 272 960 268; $lw.ZOrder(0) }

                $gwTxt = if ($sl.giant_word) { [string]$sl.giant_word } else { "STUDIO" }
                $gw = Add-Text $s "!!GiantWord" $gwTxt 30 22 900 176 $DisplayFont 124 $true $cAcc1 2 3; $gw.ZOrder(0)
                # Strict 256pt side margins so 3D object at x=320..640 never collides with text
                Add-Text $s "s2_pl" $sl.left_label 36 184 256 18 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "s2_pt" $sl.left_text 36 204 256 66 $SemiFont 12 $false $cInk 1 1 | Out-Null
                Add-Text $s "s2_sl" $sl.right_label 668 184 256 18 $MonoFont 9.5 $true $cAcc1 3 3 | Out-Null
                Add-Text $s "s2_st" $sl.right_text 668 204 256 66 $SemiFont 12 $false $cInk 3 1 | Out-Null

                $kpiXs = @(36, 192, 624, 780); $kpis = @($sl.kpis)
                for ($i = 0; $i -lt [Math]::Min(4, $kpis.Count); $i++) {
                    Add-KpiBox $s "s2_k$i" $kpis[$i].value $kpis[$i].label $kpiXs[$i] 396 144 78 $cSurface $cAcc1 $cInk $cBorder | Out-Null
                }
                $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 324 446 312 62; $sh.ZOrder(0)
                if ($heroGlb) {
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 320 68 320 356 8 -14 0 | Out-Null
                } elseif ($sl._hero_cutout) {
                    Add-Pic $s "!!HeroCut" $sl._hero_cutout 320 68 320 356 | Out-Null
                }
                Set-Morph $s ([single]2.1)
            }

            "vertical_world_elevator" {
                $stop = if ($sl.elevator_stop) { [string]$sl.elevator_stop } else { "upper_sky" }
                $gwTxt = if ($sl.giant_word) { [string]$sl.giant_word } else { "HORIZON" }
                $keep = @("!!HeroShadow", "!!Hero3D")
                if ($prevType -eq "vertical_world_elevator") { $keep += @("!!elev_bg", "!!elev_word", "!!elev_fg", "!!ev_rcard_0", "!!ev_rcard_1", "!!ev_rcard_2") }
                $s = Start-MorphSlide $pres $prevSlide $keep
                $ebgP = if (Test-Path "$AssetsDir\elev_bg.png") { "$AssetsDir\elev_bg.png" } else { "$AssetsDir\bg_base.png" }

                if ($stop -eq "lower_ground") {
                    $ebg = Add-Pic $s "!!elev_bg" $ebgP 0 -520 960 1080; $ebg.ZOrder(1)
                    $ew = Add-Text $s "!!elev_word" $gwTxt 40 290 880 170 $HeadingFont 128 $true $cWhite 2 3; Add-TextShadow $ew; $ew.ZOrder(0)
                    if (Test-Path "$AssetsDir\elev_fg.png") { $efg = Add-Pic $s "!!elev_fg" "$AssetsDir\elev_fg.png" 0 -520 960 1080; $efg.ZOrder(0) }
                    if ($heroGlb) {
                        $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 46 418 356 62; $sh.ZOrder(0)
                        Ensure-Model3D $s "!!Hero3D" $heroGlb 62 74 328 362 16 28 -3 | Out-Null
                    }
                    Add-Text $s "ev_ovr2" $sl.overline 440 34 474 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    $eh2 = Add-Text $s "ev_h2" $sl.headline 440 56 474 68 $DisplayFont 25 $true $cWhite 1 1; Add-TextShadow $eh2
                    $cards = @($sl.cards)
                    for ($ci = 0; $ci -lt [Math]::Min(3, $cards.Count); $ci++) {
                        $cy = [single](142 + $ci * 120); $crd = $cards[$ci]
                        $rc = Add-Pic $s "!!ev_rcard_$ci" "$AssetsDir\card_spec_$($ci + 1).png" 440 $cy 474 104; $rc.ZOrder(0)
                        Add-Text $s "ev_rmet_$ci" $crd.value 462 ($cy + 16) 128 68 $DisplayFont 28 $true $cAcc1 2 3 | Out-Null
                        Add-Text $s "ev_rttl_$ci" $crd.title 598 ($cy + 16) 294 26 $DisplayFont 14 $true $cInk 1 3 | Out-Null
                        Add-Text $s "ev_rdsc_$ci" $crd.desc 598 ($cy + 46) 294 44 $BodyFont 11 $false $cMuted 1 1 | Out-Null
                    }
                    Set-Morph $s ([single]2.2)
                } else {
                    $ebg = Add-Pic $s "!!elev_bg" $ebgP 0 0 960 1080; $ebg.ZOrder(1)
                    $ew = Add-Text $s "!!elev_word" $gwTxt 40 96 880 160 $HeadingFont 128 $true $cWhite 2 3; Add-TextShadow $ew; $ew.ZOrder(0)
                    if (Test-Path "$AssetsDir\elev_fg.png") { $efg = Add-Pic $s "!!elev_fg" "$AssetsDir\elev_fg.png" 0 0 960 1080; $efg.ZOrder(0) }
                    Add-Pic $s "ev_panel" "$AssetsDir\card_panel.png" 46 250 468 242 | Out-Null
                    Add-Text $s "ev_ovr" $sl.overline 72 268 416 18 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "ev_h1" $sl.headline 72 292 416 62 $DisplayFont 22 $true $cInk 1 1 | Out-Null
                    Add-Text $s "ev_body" $sl.body 72 358 416 76 $BodyFont 11.5 $false $cMuted 1 1 | Out-Null
                    if ($sl.badge) { Add-CapsulePill $s "ev_badge" $sl.badge 72 444 198 28 $cAcc1 0 $cWhite 9.5 $MonoFont | Out-Null }
                    if ($heroGlb) {
                        $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 548 412 356 62; $sh.ZOrder(0)
                        Ensure-Model3D $s "!!Hero3D" $heroGlb 564 68 326 360 10 -22 3 | Out-Null
                    }
                    if ($nextType -eq "vertical_world_elevator") {
                        for ($ci = 0; $ci -lt 3; $ci++) {
                            Add-Pic $s "!!ev_rcard_$ci" "$AssetsDir\card_spec_$($ci + 1).png" (1080 + $ci * 160) (142 + $ci * 120) 474 104 | Out-Null
                        }
                    }
                    Set-Morph $s ([single]2.1)
                }
            }

            "rotational_drift_3d" {
                $keep = @("!!HeroShadow", "!!Hero3D", "!!HeroCut", "!!md_card_0", "!!md_card_1", "!!md_card_2")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bgP = if ($sl._bg_image) { [string]$sl._bg_image } else { "$AssetsDir\bg_base.png" }
                $bg = Add-Pic $s "md_bg" $bgP 0 0 960 540; $bg.ZOrder(1)

                Add-Text $s "md_ovr" $sl.overline 46 22 520 18 $MonoFont 10 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "md_h1" $sl.headline 46 46 470 76 $DisplayFont 25 $true $cInk 1 1 | Out-Null
                if ($heroGlb) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 494 428 410 70; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 496 52 396 398 20 -36 -4 | Out-Null
                } elseif ($sl._hero_cutout) {
                    Add-Pic $s "!!HeroCut" $sl._hero_cutout 496 52 396 398 | Out-Null
                }
                $callouts = @($sl.callouts); $cCols = @($cAcc1, $cAcc2, $cAcc3)
                for ($ci = 0; $ci -lt [Math]::Min(3, $callouts.Count); $ci++) {
                    $cy = [single](142 + $ci * 116); $ac = $cCols[$ci]
                    $mc = Add-Pic $s "!!md_card_$ci" "$AssetsDir\card_spec_$($ci + 1).png" 46 $cy 384 96; $mc.ZOrder(0)
                    Add-Text $s "md_ttl_$ci" $callouts[$ci].title 68 ($cy + 14) 344 22 $DisplayFont 13 $true $ac 1 3 | Out-Null
                    Add-Text $s "md_dsc_$ci" $callouts[$ci].desc 68 ($cy + 40) 344 46 $BodyFont 11 $false $cInk 1 1 | Out-Null
                    Add-Rect $s "md_line_$ci" 430 ($cy + 46) 56 2 $ac 0 -1 0 0 | Out-Null
                }
                Set-Morph $s ([single]2.0)
            }

            "swiss_bento_morph" {
                $keep = @("!!HeroShadow", "!!Hero3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Pic $s "bg_slide" "$AssetsDir\bg_base.png" 0 0 960 540 | Out-Null
                Add-Text $s "kb_ovr" $sl.overline 48 24 620 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "kb_h1" $sl.headline 48 46 680 54 $DisplayFont 25 $true $cInk 1 1 | Out-Null

                $hero = $sl.hero_tile
                Add-Pic $s "kb_hero" "$AssetsDir\card_bento_strip.png" 46 106 588 154 | Out-Null
                Add-Text $s "kb_htag" $hero.label 74 120 180 18 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "kb_hval" $hero.value 74 144 180 84 $DisplayFont 44 $true $cInk 1 1 | Out-Null
                Add-Rect $s "kb_vdiv" 262 128 1 108 $cBorder 0 -1 0 0 | Out-Null
                Add-Text $s "kb_hsub" $hero.title 280 124 334 30 $DisplayFont 16.5 $true $cInk 1 3 | Out-Null
                Add-Text $s "kb_hdsc" $hero.desc 280 158 334 62 $BodyFont 11.5 $false $cMuted 1 1 | Out-Null

                if ($heroGlb) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 654 234 260 44; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 668 42 232 214 14 -28 3 | Out-Null
                } elseif ($sl._hero_cutout) {
                    Add-Pic $s "!!HeroCut" $sl._hero_cutout 668 42 232 214 | Out-Null
                }

                $tiles = @($sl.tiles); $bCols = @($cAcc1, $cAcc2, $cAcc3)
                for ($ti = 0; $ti -lt [Math]::Min(3, $tiles.Count); $ti++) {
                    $t = $tiles[$ti]; $tx = [single](46 + $ti * 296)
                    Add-Pic $s "kb_tbg_$ti" "$AssetsDir\card_bento_tile_$($ti + 1).png" $tx 276 276 222 | Out-Null
                    Add-Text $s "kb_tlbl_$ti" $t.label ($tx + 22) 298 232 20 $MonoFont 9.5 $true $bCols[$ti % 3] 1 3 | Out-Null
                    Add-Text $s "kb_tval_$ti" $t.value ($tx + 22) 324 232 62 $DisplayFont 38 $true $cInk 1 3 | Out-Null
                    Add-Text $s "kb_tdsc_$ti" $t.desc ($tx + 22) 396 232 82 $BodyFont 11.5 $false $cMuted 1 1 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "timeline_rail_morph" {
                $stages = @($sl.stages); $nSt = [Math]::Max(1, [Math]::Min(4, $stages.Count))
                $actIdx = [Math]::Max(0, [Math]::Min($nSt - 1, [int]$sl.active_index))
                $keep = @("!!HeroShadow", "!!Hero3D")
                if ($prevType -eq "timeline_rail_morph") { $keep += @("bg_slide", "!!tl_rail", "!!tl_progress", "!!tl_drawer"); for ($ti = 0; $ti -lt $nSt; $ti++) { $keep += "!!tl_node_$ti" } }
                $s = Start-MorphSlide $pres $prevSlide $keep
                if ($null -eq (Find-Shape $s "bg_slide")) { Add-Pic $s "bg_slide" "$AssetsDir\bg_base.png" 0 0 960 540 | Out-Null }

                Add-Text $s "tl_ovr" $sl.overline 52 24 620 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "tl_h1" $sl.headline 52 46 760 56 $DisplayFont 25 $true $cInk 1 1 | Out-Null
                Add-Rect $s "!!tl_rail" 76 144 808 4 $cBorder 0 -1 0 0 | Out-Null
                $centers = @(); for ($ti = 0; $ti -lt $nSt; $ti++) { $centers += if ($nSt -eq 1) { [single]480 } else { [single](146 + 668 * ($ti / ($nSt - 1))) } }
                Add-Rect $s "!!tl_progress" 76 144 ([Math]::Max([single]24, $centers[$actIdx] - 76)) 4 $cAcc1 0 -1 0 0 | Out-Null
                for ($ti = 0; $ti -lt $nSt; $ti++) {
                    $stg = $stages[$ti]; $cx = $centers[$ti]
                    $nw = if ($ti -eq $actIdx) { [single]172 } else { [single]148 }
                    $bgF = if ($ti -eq $actIdx) { $cAcc1 } else { $cSurface }
                    $fgT = if ($ti -eq $actIdx) { $cWhite } else { $cMuted }
                    Add-CapsulePill $s "!!tl_node_$ti" "$($stg.code)  $($stg.label)" ([single]($cx - $nw / 2)) ([single]128) $nw ([single]34) $bgF 0 $fgT 9.5 $MonoFont | Out-Null
                }
                if ($heroGlb) {
                    $rotYs = @(22, -24, 38, -16)
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 32 448 276 54; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 44 192 252 276 12 ([single]$rotYs[$actIdx % 4]) 2 | Out-Null
                }
                $cur = $stages[$actIdx]
                Add-Pic $s "!!tl_drawer" "$AssetsDir\card_panel.png" 316 204 596 292 | Out-Null
                Add-Text $s "tl_metric" $cur.metric 344 240 176 76 $DisplayFont 42 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "tl_mlbl" $cur.metric_label 344 322 176 36 $SemiFont 11 $false $cInk 1 1 | Out-Null
                Add-Text $s "tl_sh1" $cur.title 544 226 346 52 $DisplayFont 18.5 $true $cInk 1 1 | Out-Null
                Add-Text $s "tl_sum" $cur.summary 544 284 346 110 $BodyFont 12 $false $cMuted 1 1 | Out-Null
                Set-Morph $s ([single]1.8)
            }

            "chart_infographic" {
                $keep = @("!!HeroShadow", "!!Hero3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Pic $s "bg_slide" "$AssetsDir\bg_base.png" 0 0 960 540 | Out-Null
                Add-Text $s "ch_ovr" $sl.overline 46 24 600 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "ch_h1" $sl.headline 46 46 600 54 $DisplayFont 25 $true $cInk 1 1 | Out-Null
                Add-Pic $s "ch_img" "$AssetsDir\s$idx`_chart.png" 46 110 584 384 | Out-Null
                if ($heroGlb) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 656 244 248 44; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 672 38 220 222 14 -24 2 | Out-Null
                }
                $hl = $sl.highlight
                Add-Pic $s "ch_hlbg" "$AssetsDir\card_spec_1.png" 648 278 268 104 | Out-Null
                Add-Text $s "ch_hlv" "$($hl.value)  $($hl.label)" 670 290 232 30 $DisplayFont 18 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "ch_hld" $hl.desc 670 324 232 48 $BodyFont 10.5 $false $cInk 1 1 | Out-Null
                Set-Morph $s ([single]1.8)
            }

            "architecture_blueprint" {
                $keep = @("!!HeroShadow", "!!Hero3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                Add-Pic $s "bg_slide" "$AssetsDir\bg_base.png" 0 0 960 540 | Out-Null
                Add-Text $s "ar_ovr" $sl.overline 42 22 620 20 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "ar_h1" $sl.headline 42 46 630 52 $DisplayFont 24 $true $cInk 1 1 | Out-Null
                Add-Pic $s "ar_img" "$AssetsDir\s$idx`_arch.png" 42 104 624 398 | Out-Null
                if ($heroGlb) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 686 244 236 42; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 698 36 212 222 22 -34 -3 | Out-Null
                }
                $pillars = @($sl.pillars)
                for ($pi = 0; $pi -lt [Math]::Min(2, $pillars.Count); $pi++) {
                    $pl = $pillars[$pi]; $sy = [single](276 + $pi * 116)
                    Add-Pic $s "ar_cbg_$pi" "$AssetsDir\card_spec_$($pi + 1).png" 682 $sy 236 106 | Out-Null
                    Add-Text $s "ar_ctag_$pi" "$($pl.tag)  //  $($pl.title)" 702 ($sy + 12) 204 22 $DisplayFont 12 $true $cAcc1 1 3 | Out-Null
                    Add-Text $s "ar_cdsc_$pi" $pl.desc 702 ($sy + 38) 204 58 $BodyFont 10.5 $false $cInk 1 1 | Out-Null
                }
                Set-Morph $s ([single]1.8)
            }

            "closing_finale" {
                $keep = @("!!HeroShadow", "!!Hero3D")
                $s = Start-MorphSlide $pres $prevSlide $keep
                $bgP = if ($sl._bg_image) { [string]$sl._bg_image } else { "$AssetsDir\bg_base.png" }
                $bg = Add-Pic $s "bg_slide" $bgP 0 0 960 540; $bg.ZOrder(1)
                Add-Text $s "fs_ovr" $sl.overline 52 54 460 22 $MonoFont 9.5 $true $cAcc1 1 3 | Out-Null
                Add-Text $s "fs_h1" $sl.headline 52 96 500 118 $DisplayFont 36 $true $cInk 1 1 | Out-Null
                Add-Rect $s "fs_rule" 52 228 110 ([single]2.5) $cAcc1 0 -1 0 0 | Out-Null
                Add-Text $s "fs_sub" $sl.summary 52 248 490 76 $BodyFont 13.5 $false $cMuted 1 1 | Out-Null
                $metrics = @($sl.metrics)
                for ($mi = 0; $mi -lt [Math]::Min(3, $metrics.Count); $mi++) {
                    Add-KpiBox $s "fs_mk_$mi" $metrics[$mi].value $metrics[$mi].label ([single](52 + $mi * 166)) 344 152 74 $cSurface $cAcc1 $cInk $cBorder | Out-Null
                }
                if ($sl.cta) { Add-CapsulePill $s "fs_cta" $sl.cta 52 442 286 40 $cAcc1 0 $cWhite 10.5 $DisplayFont | Out-Null }
                if ($heroGlb) {
                    $sh = Add-Pic $s "!!HeroShadow" "$AssetsDir\shadow_hero.png" 536 424 352 62; $sh.ZOrder(0)
                    Ensure-Model3D $s "!!Hero3D" $heroGlb 548 66 336 370 12 -24 2 | Out-Null
                } elseif ($sl._hero_cutout) {
                    Add-Pic $s "!!HeroCut" $sl._hero_cutout 548 62 346 408 | Out-Null
                }
                Set-Morph $s ([single]2.0)
            }

            default { throw "Unsupported slide type '$st' at index $idx" }
        }

        # Pre-stage off-canvas Morph actors for the NEXT slide's signature entry
        if ($nextType -eq "odometer_counter") {
            $nextBg = if ($nextSl.bg_hex) { HexToBgr $nextSl.bg_hex } else { $cObsidian }
            Add-Rect $s "!!odo_curtain" -980 0 960 540 $nextBg 0 -1 0 0 | Out-Null
            $nextChars = ([string]$nextSl.target_value).ToCharArray()
            $totW = 0; foreach ($ch in $nextChars) { $totW += if ($ch -eq ' ') { 24 } elseif ($ch -match '\d') { 50 } else { 32 } }
            $cx = [single](480 - ($totW / 2))
            for ($ci = 0; $ci -lt $nextChars.Count; $ci++) {
                $ch = $nextChars[$ci]
                if ($ch -eq ' ') { $cx += 24; continue }
                if ($ch -match '\d') {
                    Add-OdometerStrip $s "!!odo_col_$ci" $cx ([single](580 + (($ci * 2) % 4) * 55)) 50 $cWhite $HeadingFont | Out-Null
                    $cx += 50
                } else { $cx += 32 }
            }
        } elseif ($nextType -eq "flyby_wipe") {
            Add-Pic $s "!!flyby_cloud" "$AssetsDir\flyby_cloud.png" -880 340 760 440 | Out-Null
            if ($nextSl._flyby_cutout) {
                $foPre = Add-Pic $s "!!flyby_obj" $nextSl._flyby_cutout -440 420 420 260
                $foPre.Rotation = [single]-12
            }
        } elseif ($nextType -eq "poster_depth_reveal") {
            Add-Rect $s "!!poster_band" 348 -560 264 540 $cAcc1 0 -1 0 0 | Out-Null
            if ($nextSl._hero_cutout) { Add-Pic $s "!!poster_hero" $nextSl._hero_cutout 170 590 620 464 | Out-Null }
        } elseif ($nextType -eq "rotational_drift_3d" -and $st -ne "rotational_drift_3d") {
            for ($mi = 0; $mi -lt 3; $mi++) {
                Add-Pic $s "!!md_card_$mi" "$AssetsDir\card_spec_$($mi + 1).png" (-460 - $mi * 140) (142 + $mi * 116) 384 96 | Out-Null
            }
        }

        $prevSlide = $s
        $prevType  = $st
    }

    if (Test-Path $OutputPptx) { Remove-Item $OutputPptx -Force -ErrorAction SilentlyContinue }
    $pres.SaveAs($OutputPptx)
    Write-Host ("Saved Cinema-Morph deck -> " + $OutputPptx)

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
