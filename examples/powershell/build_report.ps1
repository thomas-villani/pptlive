<#
.SYNOPSIS
    Build a small report deck through the pptlive CLI — slides, a table, a chart.

.DESCRIPTION
    Adds a title slide, a table slide (filled cell-by-cell via `cell:S:N:R:C`
    anchors), and a chart slide (data passed inline to `shape add --kind chart`),
    then prints the deck outline. Each call is one polite, single-Ctrl-Z edit.

    Like quickstart.ps1, it first ensures PowerPoint is running with a blank
    presentation, then drives everything through `pptlive`.

.NOTES
    Requires Windows + PowerPoint and `pptlive` on PATH (or prefix with `uv run`).
#>

$ErrorActionPreference = 'Stop'

function Invoke-Pptlive {
    param([Parameter(ValueFromRemainingArguments = $true)] [string[]] $PptArgs)
    $out = & pptlive @PptArgs 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "pptlive $($PptArgs -join ' ') failed (exit $LASTEXITCODE): $out"
    }
    return ($out | ConvertFrom-Json)
}

# Ensure a running PowerPoint with one blank deck.
Write-Host "Opening a blank presentation..."
$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = $true
$null = $pp.Presentations.Add()

# --- Title slide ---
$title = Invoke-Pptlive slide add --layout title
$null = Invoke-Pptlive write --anchor-id "ph:$($title.index):ctrtitle" --text "Q3 Report"
$null = Invoke-Pptlive write --anchor-id "ph:$($title.index):subtitle" --text "Acme Corp"

# --- Table slide: 4 rows x 3 cols, filled cell by cell ---
$tableSlide = Invoke-Pptlive slide add --layout title_and_content
$ts = $tableSlide.index
$null = Invoke-Pptlive write --anchor-id "ph:${ts}:title" --text "Headcount"
$tableShape = Invoke-Pptlive shape add --slide $ts --kind table --rows 4 --cols 3 `
    --left 72 --top 150 --width 480
$tn = $tableShape.index  # z-order index of the new table shape

$grid = @(
    @('Team', 'People', 'Open roles'),
    @('Engineering', '24', '3'),
    @('Sales', '11', '2'),
    @('Support', '7', '1')
)
for ($r = 0; $r -lt $grid.Count; $r++) {
    for ($c = 0; $c -lt $grid[$r].Count; $c++) {
        $anchor = "cell:${ts}:${tn}:$($r + 1):$($c + 1)"
        $null = Invoke-Pptlive write --anchor-id $anchor --text $grid[$r][$c]
    }
}
Write-Host "Filled a $($grid.Count)x$($grid[0].Count) table on slide $ts."

# --- Chart slide: data passed inline (JSON series; comma-separated categories) ---
$chartSlide = Invoke-Pptlive slide add --layout title_and_content
$cs = $chartSlide.index
$null = Invoke-Pptlive write --anchor-id "ph:${cs}:title" --text "Revenue by quarter"
$null = Invoke-Pptlive shape add --slide $cs --kind chart --chart-type column `
    --categories "Q1,Q2,Q3,Q4" `
    --series '{"Revenue":[10,14,19,23],"Profit":[3,5,8,11]}' `
    --left 72 --top 150
Write-Host "Added a column chart on slide $cs."

# --- Show the result ---
Write-Host "`nDeck outline:"
$outline = Invoke-Pptlive outline
foreach ($item in $outline) {
    Write-Host "  [$($item.slide)] $($item.title)"
}
