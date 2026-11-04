param([string[]]$Names)
# Renders _source/<name>.html to <outname>.pdf in the parent folder using headless Chrome.
$chrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$map = @{
  "01_slides" = "01_Codeflow_UIUX_Session_Slides.pdf";
  "02_srs" = "02_University_Library_System_Requirements.pdf";
  "03_wireframes" = "03_Library_System_UIUX_Wireframes.pdf"
}
$src = $PSScriptRoot
$out = Split-Path $src -Parent
foreach ($n in $Names) {
  $html = Join-Path $src "$n.html"
  $pdf = Join-Path $out $map[$n]
  $url = "file:///" + ($html -replace '\\','/' -replace ' ','%20')
  & $chrome --headless=new --disable-gpu --no-pdf-header-footer --virtual-time-budget=4000 "--print-to-pdf=$pdf" $url 2>$null | Out-Null
  Get-Item $pdf | Select-Object Name, Length, LastWriteTime
}
