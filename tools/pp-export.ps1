# pp-export.ps1 -Pptx <file> -OutDir <dir> -Json <file> [-Single] [-NoMeasure]
# Opens the pptx in PowerPoint (COM, read-only, no window), exports every slide as PNG at 2 px per pt (960x540 pt -> 1920x1080)
# (-Single: <OutDir>/render.png for slide 1; else <OutDir>/sNN.png) and writes JSON:
# { w, h (pt), slides: [ { i, overflow: [ {name, text, boxH, textH} | {name, text, axis:'x', boxW, textW} ], tables: [ {name, rows:[pt], text:[row text]} ] } ], ms }
# -NoMeasure skips the overflow/table walk (source-deck intake).
# Serialized across processes by mkdir lock %TEMP%\pptx-kit-pp.lock (stale after 180 s without heartbeat; touched per slide). Quits PowerPoint only when no presentation is open.
param(
  [Parameter(Mandatory = $true)][string]$Pptx,
  [Parameter(Mandatory = $true)][string]$OutDir,
  [Parameter(Mandatory = $true)][string]$Json,
  [switch]$Single,
  [switch]$NoMeasure
)
$ErrorActionPreference = 'Stop'
$sw = [Diagnostics.Stopwatch]::StartNew()
$lock = Join-Path ([IO.Path]::GetTempPath()) 'pptx-kit-pp.lock'

function Get-Lock {
  $deadline = (Get-Date).AddSeconds(600)
  while ($true) {
    try {
      New-Item -ItemType Directory -Path $lock -ErrorAction Stop | Out-Null
      return
    } catch {
      $info = Get-Item $lock -ErrorAction SilentlyContinue
      if ($info -and ((Get-Date) - $info.LastWriteTime).TotalSeconds -gt 180) {
        Remove-Item $lock -Recurse -Force -ErrorAction SilentlyContinue
        continue
      }
      if ((Get-Date) -gt $deadline) { throw "pp lock wait timeout" }
      Start-Sleep -Milliseconds 150
    }
  }
}

function Esc([string]$s) {
  if ($null -eq $s) { return '""' }
  $sb = New-Object Text.StringBuilder
  [void]$sb.Append('"')
  foreach ($ch in $s.ToCharArray()) {
    $c = [int]$ch
    if ($ch -eq '"') { [void]$sb.Append('\"') }
    elseif ($ch -eq '\') { [void]$sb.Append('\\') }
    elseif ($c -lt 32) { [void]$sb.Append(('\u{0:x4}' -f $c)) }
    else { [void]$sb.Append($ch) }
  }
  [void]$sb.Append('"')
  return $sb.ToString()
}
function Num($v) { return ([math]::Round([double]$v, 2)).ToString([Globalization.CultureInfo]::InvariantCulture) }
function Snip([string]$t) {
  if ($null -eq $t) { return '' }
  $t = $t -replace "[\r\n\v]+", ' / '
  if ($t.Length -gt 60) { $t = $t.Substring(0, 60) + '...' }
  return $t
}

function Walk-Shapes($shapes, $over, $tables) {
  for ($k = 1; $k -le $shapes.Count; $k++) {
    $sh = $shapes.Item($k)
    try {
      if ($sh.Type -eq 6) { Walk-Shapes $sh.GroupItems $over $tables; continue }
      if ($sh.HasTable -eq -1) {
        $tb = $sh.Table
        $rows = @(); $texts = @()
        for ($r = 1; $r -le $tb.Rows.Count; $r++) {
          $rows += (Num $tb.Rows.Item($r).Height)
          $parts = @()
          for ($c = 1; $c -le $tb.Columns.Count; $c++) {
            $parts += $tb.Cell($r, $c).Shape.TextFrame2.TextRange.Text
          }
          $texts += (Esc (Snip ($parts -join ' | ')))
        }
        [void]$tables.Add('{"name":' + (Esc $sh.Name) + ',"rows":[' + ($rows -join ',') + '],"text":[' + ($texts -join ',') + ']}')
        continue
      }
      if ($sh.HasTextFrame -eq -1) {
        $tf = $sh.TextFrame2
        if ($tf.HasText -eq -1) {
          $tr = $tf.TextRange
          $avail = $sh.Height - $tf.MarginTop - $tf.MarginBottom
          $bh = $tr.BoundHeight
          if ($bh - $avail -gt 1) {
            [void]$over.Add('{"name":' + (Esc $sh.Name) + ',"text":' + (Esc (Snip $tr.Text)) + ',"boxH":' + (Num $avail) + ',"textH":' + (Num $bh) + '}')
          }
          if ($tf.WordWrap -eq 0) {
            # wrap off: report only text spilling past the shape's outer width (insets may be overrun, as in the source deck)
            $availW = $sh.Width
            $bw = $tr.BoundWidth
            if ($bw - $availW -gt 1) {
              [void]$over.Add('{"name":' + (Esc $sh.Name) + ',"text":' + (Esc (Snip $tr.Text)) + ',"axis":"x","boxW":' + (Num $availW) + ',"textW":' + (Num $bw) + '}')
            }
          }
        }
      }
    } catch {
      [void]$over.Add('{"name":' + (Esc $sh.Name) + ',"error":' + (Esc $_.Exception.Message) + '}')
    }
  }
}

Get-Lock
$app = $null; $pres = $null
$slidesJson = New-Object Collections.ArrayList
try {
  $app = New-Object -ComObject PowerPoint.Application
  $full = (Resolve-Path $Pptx).Path
  New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
  $outFull = (Resolve-Path $OutDir).Path
  # Open(FileName, ReadOnly, Untitled, WithWindow)
  $pres = $app.Presentations.Open($full, -1, 0, 0)
  $n = $pres.Slides.Count
  $sw0 = $pres.PageSetup.SlideWidth; $sh0 = $pres.PageSetup.SlideHeight
  $pxW = [int][math]::Round($sw0 * 2); $pxH = [int][math]::Round($sh0 * 2)
  for ($i = 1; $i -le $n; $i++) {
    $slide = $pres.Slides.Item($i)
    if ($Single) { $png = Join-Path $outFull 'render.png' } else { $png = Join-Path $outFull ('s{0:D2}.png' -f $i) }
    $slide.Export($png, 'PNG', $pxW, $pxH)
    try { (Get-Item $lock).LastWriteTime = Get-Date } catch {}   # heartbeat: long decks never look stale
    $over = New-Object Collections.ArrayList
    $tables = New-Object Collections.ArrayList
    if (-not $NoMeasure) { Walk-Shapes $slide.Shapes $over $tables }
    [void]$slidesJson.Add('{"i":' + $i + ',"overflow":[' + ($over -join ',') + '],"tables":[' + ($tables -join ',') + ']}')
  }
} finally {
  if ($pres) { try { $pres.Close() } catch {} ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($app) {
    try { if ($app.Presentations.Count -eq 0) { $app.Quit() } } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Remove-Item $lock -Recurse -Force -ErrorAction SilentlyContinue
}
$out = '{"w":' + (Num $sw0) + ',"h":' + (Num $sh0) + ',"slides":[' + ($slidesJson -join ',') + '],"ms":' + $sw.ElapsedMilliseconds + '}'
[IO.File]::WriteAllText($Json, $out, (New-Object Text.UTF8Encoding $false))
