<#
.SYNOPSIS
    pptlive CLI quickstart — add a couple of slides and read them back.

.DESCRIPTION
    Drives a live PowerPoint entirely through the `pptlive` CLI. Each call prints
    exactly one JSON object on stdout, which we parse with ConvertFrom-Json.
    Mutations are polite (your view is preserved) and each is a single Ctrl-Z.

    The script first ensures PowerPoint is running with a blank presentation
    (the CLI attaches to a deck; it doesn't create one), then the `pptlive`
    commands take over.

.NOTES
    Requires Windows + PowerPoint, and `pptlive` on PATH (uv pip install pptlive).
    From the repo you can instead prefix commands with `uv run`.
#>

$ErrorActionPreference = 'Stop'

# Invoke the pptlive CLI and parse its single JSON object. Throws on a non-zero
# exit code, surfacing the CLI's stderr message.
function Invoke-Pptlive {
    param([Parameter(ValueFromRemainingArguments = $true)] [string[]] $PptArgs)
    $out = & pptlive @PptArgs 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "pptlive $($PptArgs -join ' ') failed (exit $LASTEXITCODE): $out"
    }
    return ($out | ConvertFrom-Json)
}

# Ensure a running PowerPoint with one blank deck for the demo to draw into.
Write-Host "Opening a blank presentation..."
$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = $true
$null = $pp.Presentations.Add()

# Confirm pptlive can see it.
$status = Invoke-Pptlive status
Write-Host "Attached to PowerPoint ($(@($status.decks).Count) open deck(s))."

# Add a title slide and fill its placeholders.
$title = Invoke-Pptlive slide add --layout title
$ti = $title.index
$null = Invoke-Pptlive write --anchor-id "ph:${ti}:ctrtitle" --text "pptlive"
$null = Invoke-Pptlive write --anchor-id "ph:${ti}:subtitle" --text "Driven from PowerShell"
Write-Host "Added title slide at index $ti."

# Add a content slide with a bulleted body (\n makes paragraphs).
$content = Invoke-Pptlive slide add --layout title_and_content
$ci = $content.index
$null = Invoke-Pptlive write --anchor-id "ph:${ci}:title" --text "Why pptlive"
$null = Invoke-Pptlive write --anchor-id "ph:${ci}:body" `
    --text "Talks to the app you already have open`nEvery change is one clean undo"
Write-Host "Added content slide at index $ci."

# Read a placeholder back (reads never move the user's view).
$read = Invoke-Pptlive read anchor --anchor-id "ph:${ti}:ctrtitle"
Write-Host "Title reads back as: '$($read.text)'"

$slides = @(Invoke-Pptlive slides)
Write-Host "Done. A $($slides.Count)-slide deck is open in PowerPoint."
