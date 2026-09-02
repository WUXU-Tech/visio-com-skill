param(
  [Parameter(Mandatory = $true)]
  [string]$SpecPath,
  [string]$OutputPath,
  [string]$OutputDir = ".\outputs",
  [switch]$Hidden,
  [switch]$CloseWhenDone
)

$ErrorActionPreference = "Stop"

function Get-Prop {
  param(
    [Parameter(Mandatory = $true)]$Object,
    [Parameter(Mandatory = $true)][string]$Name,
    $Default = $null
  )

  if ($null -eq $Object) {
    return $Default
  }

  $prop = $Object.PSObject.Properties[$Name]
  if ($null -eq $prop -or $null -eq $prop.Value) {
    return $Default
  }

  return $prop.Value
}

function To-Bool {
  param($Value, [bool]$Default = $false)
  if ($null -eq $Value) { return $Default }
  if ($Value -is [bool]) { return $Value }
  return [System.Convert]::ToBoolean($Value)
}

function Resolve-TaskPath {
  param([string]$Path)
  if ([System.IO.Path]::IsPathRooted($Path)) {
    return $Path
  }
  return (Join-Path (Get-Location) $Path)
}

function To-VisioColor {
  param($Color, [string]$Default = "RGB(0,0,0)")
  if ($null -eq $Color -or [string]::IsNullOrWhiteSpace([string]$Color)) {
    return $Default
  }

  $value = ([string]$Color).Trim()
  if ($value.StartsWith("RGB(", [System.StringComparison]::OrdinalIgnoreCase)) {
    return $value
  }

  if ($value.StartsWith("#")) {
    $hex = $value.Substring(1)
    if ($hex.Length -eq 6) {
      $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
      $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
      $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
      return "RGB($r,$g,$b)"
    }
  }

  return $value
}

$specFullPath = Resolve-Path -LiteralPath $SpecPath
$spec = Get-Content -LiteralPath $specFullPath -Raw -Encoding UTF8 | ConvertFrom-Json

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
  $outputFullDir = Resolve-TaskPath $OutputDir
  if (-not (Test-Path -LiteralPath $outputFullDir)) {
    New-Item -ItemType Directory -Path $outputFullDir | Out-Null
  }
  $fileName = Get-Prop -Object $spec -Name "fileName" -Default "visio_diagram.vsdx"
  if (-not ([string]$fileName).EndsWith(".vsdx", [System.StringComparison]::OrdinalIgnoreCase)) {
    $fileName = "$fileName.vsdx"
  }
  $vsdxPath = Join-Path $outputFullDir $fileName
}
else {
  $vsdxPath = Resolve-TaskPath $OutputPath
  $parent = Split-Path -Parent $vsdxPath
  if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent | Out-Null
  }
}

$logPath = [System.IO.Path]::ChangeExtension($vsdxPath, ".log")
foreach ($path in @($vsdxPath, $logPath)) {
  if (Test-Path -LiteralPath $path) {
    Remove-Item -LiteralPath $path -Force
  }
}

function Write-Step {
  param([string]$Message)
  $line = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
  Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
  Write-Output $line
}

$visio = $null
$doc = $null
$page = $null
$visible = -not $Hidden
$keepOpen = -not $CloseWhenDone

try {
  Write-Step "Starting Visio COM"
  $visio = New-Object -ComObject Visio.Application
  $visio.Visible = $visible
  $visio.AlertResponse = 1

  Write-Step "Creating document"
  $doc = $visio.Documents.Add("")
  $page = $visio.ActivePage

  $defaults = Get-Prop -Object $spec -Name "defaults" -Default ([pscustomobject]@{})
  $script:lineColor = To-VisioColor (Get-Prop -Object $defaults -Name "lineColor" -Default (Get-Prop -Object $defaults -Name "line" -Default "#1F2937"))
  $script:textColor = To-VisioColor (Get-Prop -Object $defaults -Name "textColor" -Default "#000000")
  $script:fillColor = To-VisioColor (Get-Prop -Object $defaults -Name "fill" -Default "#FFFFFF")
  $script:lineWeight = [double](Get-Prop -Object $defaults -Name "lineWeight" -Default 1.2)

  function Set-PageSize {
    param($Page, [double]$Width, [double]$Height)
    $Page.PageSheet.CellsU("PageWidth").FormulaU = "$Width in"
    $Page.PageSheet.CellsU("PageHeight").FormulaU = "$Height in"
  }

  function Set-TextStyle {
    param(
      [Parameter(Mandatory = $true)]$Shape,
      [double]$Size = 10,
      [string]$Color = $script:textColor,
      [bool]$Bold = $false,
      [double]$Angle = 0
    )

    $Shape.CellsU("Char.Size").FormulaU = "$Size pt"
    $Shape.CellsU("Char.Color").FormulaU = (To-VisioColor $Color $script:textColor)
    $Shape.CellsU("Char.Style").FormulaU = if ($Bold) { "1" } else { "0" }
    $Shape.CellsU("Para.HorzAlign").FormulaU = "1"
    $Shape.CellsU("VerticalAlign").FormulaU = "1"
    if ($Angle -ne 0) {
      $Shape.CellsU("Angle").FormulaU = "$Angle deg"
    }
  }

  function Set-LineStyle {
    param(
      [Parameter(Mandatory = $true)]$Shape,
      [bool]$Arrow = $false,
      [bool]$Dashed = $false,
      $EndArrow = $null,
      [double]$Weight = $script:lineWeight,
      [string]$Color = $script:lineColor
    )

    $Shape.CellsU("LineColor").FormulaU = (To-VisioColor $Color $script:lineColor)
    $Shape.CellsU("LineWeight").FormulaU = "$Weight pt"
    if ($null -ne $EndArrow) {
      $Shape.CellsU("EndArrow").ResultIU = [double]$EndArrow
      $Shape.CellsU("EndArrowSize").FormulaU = "3"
    }
    elseif ($Arrow) {
      $Shape.CellsU("EndArrow").ResultIU = 13
      $Shape.CellsU("EndArrowSize").FormulaU = "3"
    }
    if ($Dashed) {
      $Shape.CellsU("LinePattern").FormulaU = "2"
    }
  }

  function Add-TextBox {
    param(
      [double]$X1,
      [double]$Y1,
      [double]$X2,
      [double]$Y2,
      [string]$Text,
      [double]$Size = 10,
      [string]$Color = $script:textColor,
      [bool]$Bold = $false,
      [double]$Angle = 0
    )

    $shape = $page.DrawRectangle($X1, $Y1, $X2, $Y2)
    $shape.Text = $Text
    $shape.CellsU("FillPattern").FormulaU = "0"
    $shape.CellsU("LinePattern").FormulaU = "0"
    Set-TextStyle -Shape $shape -Size $Size -Color $Color -Bold $Bold -Angle $Angle
    return $shape
  }

  function Add-BasicShape {
    param(
      [string]$ShapeType,
      [double]$X,
      [double]$Y,
      [double]$W,
      [double]$H,
      [string]$Text = "",
      [string]$Fill = $script:fillColor,
      [string]$Line = $script:lineColor,
      [double]$Weight = $script:lineWeight,
      [double]$TextSize = 10,
      [string]$TextColor = $script:textColor,
      [bool]$Bold = $false,
      [bool]$Dashed = $false
    )

    $kind = if ([string]::IsNullOrWhiteSpace($ShapeType)) { "roundrect" } else { $ShapeType.ToLowerInvariant() }
    if ($kind -eq "ellipse" -or $kind -eq "oval") {
      $shape = $page.DrawOval($X, $Y, $X + $W, $Y + $H)
    }
    elseif ($kind -eq "diamond") {
      $cx = $X + ($W / 2)
      $cy = $Y + ($H / 2)
      $points = [double[]]@(
        $cx, ($cy + ($H / 2)),
        ($cx + ($W / 2)), $cy,
        $cx, ($cy - ($H / 2)),
        ($cx - ($W / 2)), $cy,
        $cx, ($cy + ($H / 2))
      )
      $shape = $page.DrawPolyline($points, 0)
    }
    else {
      $shape = $page.DrawRectangle($X, $Y, $X + $W, $Y + $H)
      if ($kind -eq "roundrect" -or $kind -eq "rounded" -or $kind -eq "roundedrectangle") {
        $shape.CellsU("Rounding").FormulaU = "0.08 in"
      }
    }

    $shape.Text = $Text
    $shape.CellsU("FillForegnd").FormulaU = (To-VisioColor $Fill $script:fillColor)
    $shape.CellsU("LineColor").FormulaU = (To-VisioColor $Line $script:lineColor)
    $shape.CellsU("LineWeight").FormulaU = "$Weight pt"
    if ($Dashed) {
      $shape.CellsU("LinePattern").FormulaU = "2"
    }
    Set-TextStyle -Shape $shape -Size $TextSize -Color $TextColor -Bold $Bold
    return $shape
  }

  function Add-FrameItem {
    param($Item)
    $shape = $page.DrawRectangle([double]$Item.x1, [double]$Item.y1, [double]$Item.x2, [double]$Item.y2)
    $shape.CellsU("FillPattern").FormulaU = "0"
    Set-LineStyle -Shape $shape `
      -Dashed (To-Bool (Get-Prop -Object $Item -Name "dashed" -Default $true)) `
      -Weight ([double](Get-Prop -Object $Item -Name "weight" -Default $script:lineWeight)) `
      -Color (Get-Prop -Object $Item -Name "lineColor" -Default $script:lineColor)
  }

  function Add-LineSegment {
    param(
      [double]$X1,
      [double]$Y1,
      [double]$X2,
      [double]$Y2,
      [bool]$Arrow = $false,
      [bool]$Dashed = $false,
      $EndArrow = $null,
      [double]$Weight = $script:lineWeight,
      [string]$Color = $script:lineColor
    )

    $shape = $page.DrawLine($X1, $Y1, $X2, $Y2)
    Set-LineStyle -Shape $shape -Arrow $Arrow -Dashed $Dashed -EndArrow $EndArrow -Weight $Weight -Color $Color
    return $shape
  }

  function Add-ConnectorPoints {
    param($Points, $Item)
    if ($null -eq $Points -or $Points.Count -lt 2) {
      throw "connector requires at least two points"
    }

    $arrow = To-Bool (Get-Prop -Object $Item -Name "arrow" -Default $false)
    $dashed = To-Bool (Get-Prop -Object $Item -Name "dashed" -Default $false)
    $weight = [double](Get-Prop -Object $Item -Name "weight" -Default (Get-Prop -Object $Item -Name "lineWeight" -Default $script:lineWeight))
    $color = Get-Prop -Object $Item -Name "lineColor" -Default (Get-Prop -Object $Item -Name "line" -Default $script:lineColor)
    $endArrow = Get-Prop -Object $Item -Name "endArrow" -Default $null

    for ($i = 0; $i -lt ($Points.Count - 1); $i++) {
      $start = $Points[$i]
      $end = $Points[$i + 1]
      $isLast = $i -eq ($Points.Count - 2)
      $segmentEndArrow = $null
      if ($isLast) {
        $segmentEndArrow = $endArrow
      }
      Add-LineSegment `
        -X1 ([double]$start[0]) -Y1 ([double]$start[1]) `
        -X2 ([double]$end[0]) -Y2 ([double]$end[1]) `
        -Arrow ($isLast -and $arrow) `
        -Dashed $dashed `
        -EndArrow $segmentEndArrow `
        -Weight $weight `
        -Color $color | Out-Null
    }
  }

  function Add-FoldBarsItem {
    param($Item)
    $x = [double]$Item.x
    $y = [double]$Item.y
    $barCount = [int](Get-Prop -Object $Item -Name "barCount" -Default 5)
    $segmentCount = [int](Get-Prop -Object $Item -Name "segmentCount" -Default 5)
    $barWidth = [double](Get-Prop -Object $Item -Name "barWidth" -Default 0.38)
    $segmentHeight = [double](Get-Prop -Object $Item -Name "segmentHeight" -Default 0.24)
    $gap = [double](Get-Prop -Object $Item -Name "gap" -Default 0.22)
    $trainFill = Get-Prop -Object $Item -Name "trainFill" -Default "#60A6C1"
    $validFill = Get-Prop -Object $Item -Name "validFill" -Default "#EF7118"
    $validationIndexes = Get-Prop -Object $Item -Name "validationIndexes" -Default @(0, 1, 2, 3, 4)

    for ($bar = 0; $bar -lt $barCount; $bar++) {
      $barX = $x + ($bar * ($barWidth + $gap))
      $validIndex = [int]$validationIndexes[$bar]
      for ($seg = 0; $seg -lt $segmentCount; $seg++) {
        $bottom = $y + ($seg * $segmentHeight)
        $isValid = $seg -eq $validIndex
        $fill = if ($isValid) { $validFill } else { $trainFill }
        $shape = $page.DrawRectangle($barX, $bottom, $barX + $barWidth, $bottom + $segmentHeight)
        $shape.CellsU("FillForegnd").FormulaU = (To-VisioColor $fill)
        $shape.CellsU("LineColor").FormulaU = $script:lineColor
        $shape.CellsU("LineWeight").FormulaU = "0.9 pt"
        if ($isValid) {
          $shape.Text = "V"
          Set-TextStyle -Shape $shape -Size 9 -Color "#FFFFFF" -Bold $true
        }
      }
    }
  }

  function Draw-ItemsPage {
    param($Items)
    foreach ($item in $Items) {
      $type = (Get-Prop -Object $item -Name "type" -Default "").ToString()
      switch ($type) {
        "text" {
          Add-TextBox `
            -X1 ([double]$item.x1) -Y1 ([double]$item.y1) -X2 ([double]$item.x2) -Y2 ([double]$item.y2) `
            -Text (Get-Prop -Object $item -Name "text" -Default "") `
            -Size ([double](Get-Prop -Object $item -Name "size" -Default 10)) `
            -Color (Get-Prop -Object $item -Name "color" -Default $script:textColor) `
            -Bold (To-Bool (Get-Prop -Object $item -Name "bold" -Default $false)) `
            -Angle ([double](Get-Prop -Object $item -Name "angle" -Default 0)) | Out-Null
        }
        "box" {
          Add-BasicShape `
            -ShapeType (Get-Prop -Object $item -Name "shape" -Default "rectangle") `
            -X ([double]$item.x1) -Y ([double]$item.y1) -W ([double]($item.x2 - $item.x1)) -H ([double]($item.y2 - $item.y1)) `
            -Text (Get-Prop -Object $item -Name "text" -Default "") `
            -Fill (Get-Prop -Object $item -Name "fill" -Default $script:fillColor) `
            -Line (Get-Prop -Object $item -Name "lineColor" -Default $script:lineColor) `
            -TextSize ([double](Get-Prop -Object $item -Name "size" -Default 10)) `
            -TextColor (Get-Prop -Object $item -Name "color" -Default $script:textColor) `
            -Bold (To-Bool (Get-Prop -Object $item -Name "bold" -Default $false)) | Out-Null
        }
        "ellipse" {
          Add-BasicShape -ShapeType "ellipse" -X ([double]$item.x1) -Y ([double]$item.y1) -W ([double]($item.x2 - $item.x1)) -H ([double]($item.y2 - $item.y1)) -Text (Get-Prop -Object $item -Name "text" -Default "") -Fill (Get-Prop -Object $item -Name "fill" -Default $script:fillColor) | Out-Null
        }
        "diamond" {
          Add-BasicShape `
            -ShapeType "diamond" `
            -X ([double]($item.cx - ($item.w / 2))) -Y ([double]($item.cy - ($item.h / 2))) -W ([double]$item.w) -H ([double]$item.h) `
            -Text (Get-Prop -Object $item -Name "text" -Default "") `
            -Fill (Get-Prop -Object $item -Name "fill" -Default "#FFF6D7") `
            -TextSize ([double](Get-Prop -Object $item -Name "size" -Default 10)) | Out-Null
        }
        "frame" { Add-FrameItem -Item $item }
        "line" {
          Add-LineSegment `
            -X1 ([double]$item.x1) -Y1 ([double]$item.y1) -X2 ([double]$item.x2) -Y2 ([double]$item.y2) `
            -Arrow (To-Bool (Get-Prop -Object $item -Name "arrow" -Default $false)) `
            -Dashed (To-Bool (Get-Prop -Object $item -Name "dashed" -Default $false)) `
            -EndArrow (Get-Prop -Object $item -Name "endArrow" -Default $null) `
            -Weight ([double](Get-Prop -Object $item -Name "weight" -Default $script:lineWeight)) `
            -Color (Get-Prop -Object $item -Name "lineColor" -Default (Get-Prop -Object $item -Name "line" -Default $script:lineColor)) | Out-Null
        }
        "connector" { Add-ConnectorPoints -Points (Get-Prop -Object $item -Name "points" -Default @()) -Item $item }
        "foldBars" { Add-FoldBarsItem -Item $item }
        default { throw "Unknown item type: $type" }
      }
    }
  }

  function Get-NodeEndpoint {
    param($From, $To)
    $fx = [double]$From.x + ([double]$From.w / 2)
    $fy = [double]$From.y + ([double]$From.h / 2)
    $tx = [double]$To.x + ([double]$To.w / 2)
    $ty = [double]$To.y + ([double]$To.h / 2)
    $dx = $tx - $fx
    $dy = $ty - $fy

    if ([Math]::Abs($dx) -ge [Math]::Abs($dy)) {
      if ($dx -ge 0) {
        return @(
          @(([double]$From.x + [double]$From.w), $fy),
          @([double]$To.x, $ty)
        )
      }
      return @(
        @([double]$From.x, $fy),
        @(([double]$To.x + [double]$To.w), $ty)
      )
    }

    if ($dy -ge 0) {
      return @(
        @($fx, ([double]$From.y + [double]$From.h)),
        @($tx, [double]$To.y)
      )
    }

    return @(
      @($fx, [double]$From.y),
      @($tx, ([double]$To.y + [double]$To.h))
    )
  }

  function Draw-StandardPage {
    param($PageSpec)
    $nodeMap = @{}

    foreach ($node in (Get-Prop -Object $PageSpec -Name "nodes" -Default @())) {
      $x = [double](Get-Prop -Object $node -Name "x" -Default 1)
      $y = [double](Get-Prop -Object $node -Name "y" -Default 1)
      $w = [double](Get-Prop -Object $node -Name "w" -Default 2)
      $h = [double](Get-Prop -Object $node -Name "h" -Default 0.8)
      $shapeType = Get-Prop -Object $node -Name "shape" -Default "roundrect"

      if ($null -ne (Get-Prop -Object $node -Name "stencil" -Default $null)) {
        $shapeType = "roundrect"
      }

      $shape = Add-BasicShape `
        -ShapeType $shapeType `
        -X $x -Y $y -W $w -H $h `
        -Text (Get-Prop -Object $node -Name "text" -Default "") `
        -Fill (Get-Prop -Object $node -Name "fill" -Default $script:fillColor) `
        -Line (Get-Prop -Object $node -Name "line" -Default $script:lineColor) `
        -Weight ([double](Get-Prop -Object $node -Name "lineWeight" -Default $script:lineWeight)) `
        -TextSize ([double](Get-Prop -Object $node -Name "fontSize" -Default 11)) `
        -TextColor (Get-Prop -Object $node -Name "fontColor" -Default $script:textColor) `
        -Bold (To-Bool (Get-Prop -Object $node -Name "bold" -Default $false))

      $id = Get-Prop -Object $node -Name "id" -Default $null
      if ($null -ne $id) {
        $nodeMap[[string]$id] = [pscustomobject]@{ shape = $shape; x = $x; y = $y; w = $w; h = $h }
      }
    }

    foreach ($line in (Get-Prop -Object $PageSpec -Name "lines" -Default @())) {
      Add-LineSegment `
        -X1 ([double]$line.x1) -Y1 ([double]$line.y1) -X2 ([double]$line.x2) -Y2 ([double]$line.y2) `
        -Arrow ($null -ne (Get-Prop -Object $line -Name "endArrow" -Default $null)) `
        -Dashed (To-Bool (Get-Prop -Object $line -Name "dashed" -Default $false)) `
        -EndArrow (Get-Prop -Object $line -Name "endArrow" -Default $null) `
        -Weight ([double](Get-Prop -Object $line -Name "lineWeight" -Default $script:lineWeight)) `
        -Color (Get-Prop -Object $line -Name "line" -Default $script:lineColor) | Out-Null
    }

    foreach ($conn in (Get-Prop -Object $PageSpec -Name "connections" -Default @())) {
      $fromId = [string](Get-Prop -Object $conn -Name "from" -Default "")
      $toId = [string](Get-Prop -Object $conn -Name "to" -Default "")
      if (-not $nodeMap.ContainsKey($fromId) -or -not $nodeMap.ContainsKey($toId)) {
        Write-Step "Skipping connection with missing endpoint: $fromId -> $toId"
        continue
      }

      $fromNode = $nodeMap[$fromId]
      $toNode = $nodeMap[$toId]
      $pair = Get-NodeEndpoint -From $fromNode -To $toNode
      $start = $pair[0]
      $end = $pair[1]
      $endArrow = Get-Prop -Object $conn -Name "endArrow" -Default 13
      Add-LineSegment `
        -X1 ([double]$start[0]) -Y1 ([double]$start[1]) -X2 ([double]$end[0]) -Y2 ([double]$end[1]) `
        -EndArrow $endArrow `
        -Dashed (To-Bool (Get-Prop -Object $conn -Name "dashed" -Default $false)) `
        -Weight ([double](Get-Prop -Object $conn -Name "lineWeight" -Default $script:lineWeight)) `
        -Color (Get-Prop -Object $conn -Name "line" -Default "#475569") | Out-Null

      $label = Get-Prop -Object $conn -Name "text" -Default ""
      if (-not [string]::IsNullOrWhiteSpace([string]$label)) {
        $pin = [double](Get-Prop -Object $conn -Name "textPinX" -Default 0.5)
        $offsetX = [double](Get-Prop -Object $conn -Name "textOffsetX" -Default 0)
        $offsetY = [double](Get-Prop -Object $conn -Name "textOffsetY" -Default 0.18)
        $lx = ([double]$start[0]) + ((([double]$end[0]) - ([double]$start[0])) * $pin) + $offsetX
        $ly = ([double]$start[1]) + ((([double]$end[1]) - ([double]$start[1])) * $pin) + $offsetY
        $labelW = [Math]::Max(0.7, [Math]::Min(2.4, ([string]$label).Length * 0.08))
        Add-TextBox -X1 ($lx - ($labelW / 2)) -Y1 ($ly - 0.12) -X2 ($lx + ($labelW / 2)) -Y2 ($ly + 0.16) -Text $label -Size ([double](Get-Prop -Object $conn -Name "fontSize" -Default 9)) | Out-Null
      }
    }

    foreach ($label in (Get-Prop -Object $PageSpec -Name "labels" -Default @())) {
      Add-TextBox `
        -X1 ([double]$label.x) -Y1 ([double]$label.y) -X2 ([double]($label.x + $label.w)) -Y2 ([double]($label.y + $label.h)) `
        -Text (Get-Prop -Object $label -Name "text" -Default "") `
        -Size ([double](Get-Prop -Object $label -Name "fontSize" -Default 10)) `
        -Color (Get-Prop -Object $label -Name "fontColor" -Default $script:textColor) `
        -Bold (To-Bool (Get-Prop -Object $label -Name "bold" -Default $false)) | Out-Null
    }
  }

  function Draw-SequenceDiagram {
    param($Spec)
    $actors = Get-Prop -Object $Spec -Name "actors" -Default @()
    $messages = Get-Prop -Object $Spec -Name "messages" -Default @()
    $layout = Get-Prop -Object $Spec -Name "layout" -Default ([pscustomobject]@{})
    $spacing = [double](Get-Prop -Object $layout -Name "actorSpacing" -Default 2.6)
    $messageSpacing = [double](Get-Prop -Object $layout -Name "messageSpacing" -Default 0.62)
    $startY = [double](Get-Prop -Object $layout -Name "startY" -Default 7.6)
    $lifelineHeight = [double](Get-Prop -Object $layout -Name "lifelineHeight" -Default ([Math]::Max(4.5, $messages.Count * $messageSpacing + 1.2)))
    $left = 1.2
    $actorW = 1.55
    $actorH = 0.55
    $pageW = [Math]::Max(8.5, $left + (($actors.Count - 1) * $spacing) + 2.2)
    $pageH = [Math]::Max(6.5, $startY + 0.8)
    Set-PageSize -Page $page -Width ([double](Get-Prop -Object $Spec -Name "pageWidth" -Default $pageW)) -Height ([double](Get-Prop -Object $Spec -Name "pageHeight" -Default $pageH))

    $title = Get-Prop -Object $Spec -Name "title" -Default ""
    if (-not [string]::IsNullOrWhiteSpace([string]$title)) {
      Add-TextBox -X1 0.7 -Y1 ($pageH - 0.55) -X2 ($pageW - 0.7) -Y2 ($pageH - 0.15) -Text $title -Size 16 -Bold $true -Color "#111827" | Out-Null
    }

    $actorMap = @{}
    for ($i = 0; $i -lt $actors.Count; $i++) {
      $actor = $actors[$i]
      $cx = $left + ($i * $spacing)
      $atype = [string](Get-Prop -Object $actor -Name "type" -Default "actor")
      $fill = switch ($atype) {
        "system" { "#F5F3FF" }
        "database" { "#FEF3C7" }
        "external" { "#ECFDF5" }
        default { "#EFF6FF" }
      }
      $line = switch ($atype) {
        "system" { "#8B5CF6" }
        "database" { "#D97706" }
        "external" { "#10B981" }
        default { "#3B82F6" }
      }

      $name = Get-Prop -Object $actor -Name "name" -Default (Get-Prop -Object $actor -Name "id" -Default "Actor")
      Add-BasicShape -ShapeType "roundrect" -X ($cx - ($actorW / 2)) -Y $startY -W $actorW -H $actorH -Text $name -Fill (Get-Prop -Object $actor -Name "fill" -Default $fill) -Line (Get-Prop -Object $actor -Name "line" -Default $line) -TextSize 10 | Out-Null
      Add-LineSegment -X1 $cx -Y1 $startY -X2 $cx -Y2 ($startY - $lifelineHeight) -Dashed $true -Weight 0.9 -Color "#64748B" | Out-Null
      $actorMap[[string](Get-Prop -Object $actor -Name "id" -Default $i)] = [pscustomobject]@{ x = $cx; y = $startY }
    }

    $y = $startY - 0.65
    foreach ($msg in $messages) {
      $fromId = [string](Get-Prop -Object $msg -Name "from" -Default "")
      $toId = [string](Get-Prop -Object $msg -Name "to" -Default "")
      if (-not $actorMap.ContainsKey($fromId) -or -not $actorMap.ContainsKey($toId)) {
        Write-Step "Skipping message with missing actor: $fromId -> $toId"
        continue
      }

      $from = $actorMap[$fromId]
      $to = $actorMap[$toId]
      $msgType = [string](Get-Prop -Object $msg -Name "type" -Default "sync")
      $isReturn = $msgType -eq "return"
      $messageArrow = 13
      if ($isReturn) {
        $messageArrow = 1
      }
      Add-LineSegment -X1 ([double]$from.x) -Y1 $y -X2 ([double]$to.x) -Y2 $y -EndArrow $messageArrow -Dashed $isReturn -Weight 1.0 -Color "#475569" | Out-Null

      $label = Get-Prop -Object $msg -Name "text" -Default ""
      if (-not [string]::IsNullOrWhiteSpace([string]$label)) {
        $lx = (([double]$from.x) + ([double]$to.x)) / 2
        Add-TextBox -X1 ($lx - 0.9) -Y1 ($y + 0.06) -X2 ($lx + 0.9) -Y2 ($y + 0.28) -Text $label -Size 8.5 | Out-Null
      }

      $activation = To-Bool (Get-Prop -Object $msg -Name "activation" -Default (-not $isReturn))
      if ($activation) {
        $ax = [double]$to.x
        Add-BasicShape -ShapeType "rectangle" -X ($ax - 0.05) -Y ($y - 0.22) -W 0.1 -H 0.36 -Text "" -Fill "#FFFFFF" -Line "#64748B" -Weight 0.8 | Out-Null
      }

      $y -= $messageSpacing
    }
  }

  function Prepare-Page {
    param($PageSpec, [int]$Index)
    if ($Index -eq 0) {
      $script:page = $doc.Pages.Item(1)
    }
    else {
      $script:page = $doc.Pages.Add()
    }
    $script:page.Name = Get-Prop -Object $PageSpec -Name "name" -Default "Page-$($Index + 1)"
    $width = [double](Get-Prop -Object $PageSpec -Name "pageWidth" -Default (Get-Prop -Object $spec -Name "pageWidth" -Default 11))
    $height = [double](Get-Prop -Object $PageSpec -Name "pageHeight" -Default (Get-Prop -Object $spec -Name "pageHeight" -Default 8.5))
    Set-PageSize -Page $script:page -Width $width -Height $height
  }

  Write-Step "Drawing diagram"

  if ((Get-Prop -Object $spec -Name "type" -Default "") -eq "sequence") {
    $page = $doc.Pages.Item(1)
    $page.Name = Get-Prop -Object $spec -Name "pageName" -Default "Sequence"
    Draw-SequenceDiagram -Spec $spec
  }
  elseif ($null -ne (Get-Prop -Object $spec -Name "items" -Default $null)) {
    $page = $doc.Pages.Item(1)
    $page.Name = Get-Prop -Object $spec -Name "pageName" -Default "Diagram"
    $pageSpec = Get-Prop -Object $spec -Name "page" -Default ([pscustomobject]@{})
    Set-PageSize -Page $page -Width ([double](Get-Prop -Object $pageSpec -Name "width" -Default 11)) -Height ([double](Get-Prop -Object $pageSpec -Name "height" -Default 8.5))
    Draw-ItemsPage -Items $spec.items
  }
  else {
    $pages = Get-Prop -Object $spec -Name "pages" -Default $null
    if ($null -eq $pages) {
      $pages = @($spec)
    }
    for ($p = 0; $p -lt $pages.Count; $p++) {
      Prepare-Page -PageSpec $pages[$p] -Index $p
      Draw-StandardPage -PageSpec $pages[$p]
    }
  }

  Write-Step "Saving VSDX"
  $doc.SaveAs($vsdxPath)
  Write-Step "Done"

  [pscustomobject]@{
    Vsdx = $vsdxPath
    Log = $logPath
  } | ConvertTo-Json
}
catch {
  Write-Step "ERROR: $($_.Exception.Message)"
  throw
}
finally {
  if ($null -ne $doc) {
    $doc.Saved = $true
  }

  if (-not ($visible -and $keepOpen)) {
    Write-Step "Closing Visio COM"
    if ($null -ne $doc) {
      $doc.Close()
    }
    if ($null -ne $visio) {
      $visio.Quit()
    }
  }
  else {
    Write-Step "Leaving Visio open for review"
  }
}
