param(
  [string]$HtmlPath = "index.html",
  [string]$OutputPath = "SHAPE-Lab-Website-Content.xlsx"
)

$ErrorActionPreference = "Stop"

function Decode-Html([string]$Value) {
  if ($null -eq $Value) { return "" }
  return [System.Net.WebUtility]::HtmlDecode($Value).Trim()
}

function Strip-Html([string]$Value) {
  if ($null -eq $Value) { return "" }
  $text = $Value -replace "(?i)<br\s*/?>", " | "
  $text = $text -replace "(?s)<script.*?</script>", " "
  $text = $text -replace "(?s)<style.*?</style>", " "
  $text = $text -replace "<[^>]+>", " "
  $text = Decode-Html($text)
  return ($text -replace "\s+", " ").Trim()
}

function Get-Attr([string]$Html, [string]$Name) {
  $match = [regex]::Match($Html, "\b$Name\s*=\s*[""']([^""']*)[""']", "IgnoreCase")
  if ($match.Success) { return Decode-Html($match.Groups[1].Value) }
  return ""
}

function Get-FirstText([string]$Html, [string]$Pattern) {
  $match = [regex]::Match($Html, $Pattern, "Singleline,IgnoreCase")
  if ($match.Success) { return Strip-Html($match.Groups[1].Value) }
  return ""
}

function Get-FirstAttr([string]$Html, [string]$Pattern, [string]$Attr) {
  $match = [regex]::Match($Html, $Pattern, "Singleline,IgnoreCase")
  if ($match.Success) { return Get-Attr $match.Value $Attr }
  return ""
}

function Get-LineNumber([string]$Text, [int]$Index) {
  return (($Text.Substring(0, $Index) -split "`n").Count)
}

function Get-Between([string]$Text, [string]$StartNeedle, [string]$EndNeedle) {
  $start = $Text.IndexOf($StartNeedle, [StringComparison]::OrdinalIgnoreCase)
  if ($start -lt 0) { return "" }
  $end = $Text.IndexOf($EndNeedle, $start + $StartNeedle.Length, [StringComparison]::OrdinalIgnoreCase)
  if ($end -lt 0) { $end = $Text.Length }
  return $Text.Substring($start, $end - $start)
}

function New-Row([object[]]$Values) {
  return [pscustomobject]@{ Values = $Values }
}

function Add-SheetXml([System.IO.Compression.ZipArchive]$Zip, [string]$Path, [object[]]$Headers, [object[]]$Rows) {
  $entry = $Zip.CreateEntry($Path)
  $stream = $entry.Open()
  $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
  $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
  $writer.Write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>')
  $allRows = @((New-Row $Headers)) + $Rows
  for ($r = 0; $r -lt $allRows.Count; $r++) {
    $rowNumber = $r + 1
    $writer.Write("<row r=""$rowNumber"">")
    $values = $allRows[$r].Values
    for ($c = 0; $c -lt $values.Count; $c++) {
      $cellRef = "$(Convert-ToColumnName ($c + 1))$rowNumber"
      $value = [string]$values[$c]
      $escaped = [System.Security.SecurityElement]::Escape($value)
      $writer.Write("<c r=""$cellRef"" t=""inlineStr""><is><t xml:space=""preserve"">$escaped</t></is></c>")
    }
    $writer.Write('</row>')
  }
  $writer.Write('</sheetData><autoFilter ref="A1:Z1"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
  $writer.Dispose()
  $stream.Dispose()
}

function Convert-ToColumnName([int]$Index) {
  $name = ""
  while ($Index -gt 0) {
    $Index--
    $name = [char](65 + ($Index % 26)) + $name
    $Index = [math]::Floor($Index / 26)
  }
  return $name
}

function Add-ZipText([System.IO.Compression.ZipArchive]$Zip, [string]$Path, [string]$Text) {
  $entry = $Zip.CreateEntry($Path)
  $stream = $entry.Open()
  $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
  $writer.Write($Text)
  $writer.Dispose()
  $stream.Dispose()
}

$htmlFullPath = Resolve-Path $HtmlPath
$html = Get-Content -LiteralPath $htmlFullPath -Raw

$siteRows = New-Object System.Collections.Generic.List[object]
$navMatches = [regex]::Matches($html, '<nav[^>]*class="nav-links"[\s\S]*?</nav>', "Singleline,IgnoreCase")
if ($navMatches.Count -gt 0) {
  foreach ($link in [regex]::Matches($navMatches[0].Value, '<a\b[^>]*href=["'']([^"'']*)["''][^>]*>([\s\S]*?)</a>', "Singleline,IgnoreCase")) {
    $siteRows.Add((New-Row @("Navigation", "Link", $(Strip-Html($link.Groups[2].Value)), $(Decode-Html($link.Groups[1].Value)), "", "", $(Get-LineNumber $html $link.Index))))
  }
}
$siteRows.Add((New-Row @("Hero", "Title", "Main title", $(Get-FirstText $html '<section id="home"[\s\S]*?<h1>([\s\S]*?)</h1>'), "", "", "")))
$siteRows.Add((New-Row @("Hero", "Subtitle", "Main subtitle", $(Get-FirstText $html '<section id="home"[\s\S]*?<p>([\s\S]*?)</p>'), "", "", "")))
foreach ($missionItem in [regex]::Matches((Get-Between $html '<section id="mission"' '<section id="research"'), '<li>[\s\S]*?<span>([\s\S]*?)</span>[\s\S]*?</li>', "Singleline,IgnoreCase")) {
  $siteRows.Add((New-Row @("Mission", "Mission Item", "", $(Strip-Html($missionItem.Groups[1].Value)), "", "", $(Get-LineNumber $html $missionItem.Index))))
}
foreach ($section in [regex]::Matches($html, '<section id=["'']([^"'']+)["''][\s\S]*?(?=<section id=|</main>)', "Singleline,IgnoreCase")) {
  $sectionId = Decode-Html($section.Groups[1].Value)
  $title = Get-FirstText $section.Value '<h2[^>]*class=["'']section-title["''][^>]*>([\s\S]*?)</h2>'
  $lede = Get-FirstText $section.Value '<p[^>]*class=["'']section-lede["''][^>]*>([\s\S]*?)</p>'
  if ($title) { $siteRows.Add((New-Row @($sectionId, "Section Title", "", $title, "", "", $(Get-LineNumber $html $section.Index)))) }
  if ($lede) { $siteRows.Add((New-Row @($sectionId, "Section Lede", "", $lede, "", "", $(Get-LineNumber $html $section.Index)))) }
}

$researchRows = New-Object System.Collections.Generic.List[object]
$research = Get-Between $html '<section id="research"' '<section id="team"'
foreach ($group in [regex]::Matches($research, '<div class="project-group reveal">([\s\S]*?)(?=<div class="project-group reveal">|</section>)', "Singleline,IgnoreCase")) {
  $groupName = Get-FirstText $group.Value '<h3>([\s\S]*?)</h3>'
  foreach ($card in [regex]::Matches($group.Value, '<article class="card">([\s\S]*?)</article>', "Singleline,IgnoreCase")) {
    $img = [regex]::Match($card.Value, '<img\b[^>]*>', "Singleline,IgnoreCase")
    $researchRows.Add((New-Row @(
      $groupName,
      $(Get-FirstText $card.Value '<div class="card-kicker">([\s\S]*?)</div>'),
      $(Get-FirstText $card.Value '<h4>([\s\S]*?)</h4>'),
      $(Get-FirstText $card.Value '<p>([\s\S]*?)</p>'),
      $(if ($img.Success) { Get-Attr $img.Value "src" } else { "" }),
      $(if ($img.Success) { Get-Attr $img.Value "alt" } else { "" }),
      $(Get-LineNumber $html ($html.IndexOf($card.Value, [StringComparison]::Ordinal)))
    )))
  }
}

$teamRows = New-Object System.Collections.Generic.List[object]
$team = Get-Between $html '<section id="team"' '<section id="publications"'
$director = [regex]::Match($team, '<article class="director reveal">([\s\S]*?)</article>', "Singleline,IgnoreCase")
if ($director.Success) {
  $img = [regex]::Match($director.Value, '<img\b[^>]*>', "Singleline,IgnoreCase")
  $paragraphs = [regex]::Matches($director.Value, '<p[^>]*>([\s\S]*?)</p>', "Singleline,IgnoreCase")
  $teamRows.Add((New-Row @(
    "Director",
    $(Get-FirstText $director.Value '<h3>([\s\S]*?)</h3>'),
    $(if ($paragraphs.Count -gt 0) { Strip-Html($paragraphs[0].Groups[1].Value) } else { "" }),
    $(if ($paragraphs.Count -gt 1) { Strip-Html($paragraphs[1].Groups[1].Value) } else { "" }),
    $(if ($paragraphs.Count -gt 2) { Strip-Html($paragraphs[2].Groups[1].Value) } else { "" }),
    $(if ($img.Success) { Get-Attr $img.Value "src" } else { "" }),
    $(if ($img.Success) { Get-Attr $img.Value "alt" } else { "" })
  )))
}
foreach ($group in [regex]::Matches($team, '<div class="team-group reveal">([\s\S]*?)(?=<div class="team-group reveal">|</section>)', "Singleline,IgnoreCase")) {
  $groupName = Get-FirstText $group.Value '<h3[^>]*>([\s\S]*?)</h3>'
  foreach ($card in [regex]::Matches($group.Value, '<article class="card profile-card">([\s\S]*?)</article>', "Singleline,IgnoreCase")) {
    $img = [regex]::Match($card.Value, '<img\b[^>]*>', "Singleline,IgnoreCase")
    $teamRows.Add((New-Row @(
      $groupName,
      $(Get-FirstText $card.Value '<h4>([\s\S]*?)</h4>'),
      "",
      "",
      "",
      $(if ($img.Success) { Get-Attr $img.Value "src" } else { "" }),
      $(if ($img.Success) { Get-Attr $img.Value "alt" } else { "" })
    )))
  }
}

$publicationRows = New-Object System.Collections.Generic.List[object]
$publications = Get-Between $html '<section id="publications"' '<section id="news"'
foreach ($group in [regex]::Matches($publications, '<section class="publication-group"[\s\S]*?<h3[^>]*>([\s\S]*?)</h3>([\s\S]*?)</section>', "Singleline,IgnoreCase")) {
  $groupName = Strip-Html($group.Groups[1].Value)
  foreach ($item in [regex]::Matches($group.Groups[2].Value, '<li class="publication-item">([\s\S]*?)</li>', "Singleline,IgnoreCase")) {
    $doi = Get-FirstAttr $item.Value '<a\b[^>]*class=["'']text-link["''][^>]*>' "href"
    $publicationRows.Add((New-Row @(
      $groupName,
      $(Get-FirstText $item.Value '<h4>([\s\S]*?)</h4>'),
      $(Get-FirstText $item.Value '<p>([\s\S]*?)</p>'),
      $(Get-FirstText $item.Value '<div class="publication-meta">([\s\S]*?)</div>'),
      $doi,
      $(Get-LineNumber $html ($html.IndexOf($item.Value, [StringComparison]::Ordinal)))
    )))
  }
}

$newsRows = New-Object System.Collections.Generic.List[object]
$news = Get-Between $html '<section id="news"' '<section id="sponsors"'
foreach ($item in [regex]::Matches($news, '<article class="news-feature reveal">([\s\S]*?)</article>', "Singleline,IgnoreCase")) {
  $img = [regex]::Match($item.Value, '<img\b[^>]*>', "Singleline,IgnoreCase")
  $newsRows.Add((New-Row @(
    $(Get-FirstText $item.Value '<h3>([\s\S]*?)</h3>'),
    $(Get-FirstText $item.Value '<p>([\s\S]*?)</p>'),
    $(if ($img.Success) { Get-Attr $img.Value "src" } else { "" }),
    $(if ($img.Success) { Get-Attr $img.Value "alt" } else { "" }),
    $(Get-LineNumber $html ($html.IndexOf($item.Value, [StringComparison]::Ordinal)))
  )))
}

$sponsorRows = New-Object System.Collections.Generic.List[object]
$sponsors = Get-Between $html '<section id="sponsors"' '<section id="contact"'
foreach ($grid in [regex]::Matches($sponsors, '<div class="sponsor-grid reveal"[^>]*aria-label=["'']([^"'']*)["''][^>]*>([\s\S]*?)</div>\s*(?=<h3|</div>\s*</section>)', "Singleline,IgnoreCase")) {
  $category = Decode-Html($grid.Groups[1].Value)
  foreach ($card in [regex]::Matches($grid.Groups[2].Value, '<a class="sponsor-card"[\s\S]*?</a>', "Singleline,IgnoreCase")) {
    $img = [regex]::Match($card.Value, '<img\b[^>]*>', "Singleline,IgnoreCase")
    $sponsorRows.Add((New-Row @(
      $category,
      $(Get-FirstText $card.Value '<h4>([\s\S]*?)</h4>'),
      $(Get-FirstText $card.Value '<div class="sponsor-label">([\s\S]*?)</div>'),
      $(Get-Attr $card.Value "href"),
      $(if ($img.Success) { Get-Attr $img.Value "src" } else { "" }),
      $(if ($img.Success) { Get-Attr $img.Value "alt" } else { "" }),
      $(Get-LineNumber $html ($html.IndexOf($card.Value, [StringComparison]::Ordinal)))
    )))
  }
}

$contactRows = New-Object System.Collections.Generic.List[object]
$contact = Get-Between $html '<section id="contact"' '</main>'
$contactRows.Add((New-Row @("Headline", "", $(Get-FirstText $contact '<h3>([\s\S]*?)</h3>'))))
$contactRows.Add((New-Row @("Description", "", $(Get-FirstText $contact '<p[^>]*class=["'']section-lede["''][^>]*>([\s\S]*?)</p>'))))
foreach ($li in [regex]::Matches($contact, '<li>([\s\S]*?)</li>', "Singleline,IgnoreCase")) {
  $text = Strip-Html($li.Groups[1].Value)
  $parts = $text -split ":", 2
  $contactRows.Add((New-Row @($(if ($parts.Count -gt 0) { $parts[0].Trim() } else { "" }), "", $(if ($parts.Count -gt 1) { $parts[1].Trim() } else { $text }))))
}

$imageRows = New-Object System.Collections.Generic.List[object]
foreach ($img in [regex]::Matches($html, '<img\b[^>]*>', "Singleline,IgnoreCase")) {
  $src = Get-Attr $img.Value "src"
  $imageRows.Add((New-Row @(
    $src,
    $(Get-Attr $img.Value "alt"),
    $(if ($src -match '^https?://') { "Remote" } else { "Local" }),
    $(Get-Attr $img.Value "loading"),
    $(Get-LineNumber $html $img.Index)
  )))
}

$linkRows = New-Object System.Collections.Generic.List[object]
foreach ($link in [regex]::Matches($html, '<a\b[^>]*href=["'']([^"'']*)["''][^>]*>([\s\S]*?)</a>', "Singleline,IgnoreCase")) {
  $linkRows.Add((New-Row @(
    $(Strip-Html($link.Groups[2].Value)),
    $(Decode-Html($link.Groups[1].Value)),
    $(Get-Attr $link.Value "aria-label"),
    $(Get-LineNumber $html $link.Index)
  )))
}

$photoRows = New-Object System.Collections.Generic.List[object]
$photoRoot = Join-Path (Split-Path $htmlFullPath) "shapelab_photos"
if (Test-Path $photoRoot) {
  foreach ($file in Get-ChildItem -LiteralPath $photoRoot -Recurse -File | Sort-Object FullName) {
    $photoRows.Add((New-Row @(
      $file.FullName.Substring((Split-Path $htmlFullPath).Length + 1),
      $file.Extension.TrimStart("."),
      $file.Length,
      $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
    )))
  }
}

$sheets = @(
  @{ Name = "Site Copy"; Headers = @("Section", "Content Type", "Label", "Text Or Value", "Link", "Notes", "Source Line"); Rows = $siteRows.ToArray() },
  @{ Name = "Research"; Headers = @("Group", "Kicker", "Title", "Description", "Image Src", "Image Alt", "Source Line"); Rows = $researchRows.ToArray() },
  @{ Name = "Team"; Headers = @("Group", "Name", "Role", "Education", "Focus", "Image Src", "Image Alt"); Rows = $teamRows.ToArray() },
  @{ Name = "Publications"; Headers = @("Group", "Title", "Authors", "Publication Meta", "DOI Or Link", "Source Line"); Rows = $publicationRows.ToArray() },
  @{ Name = "News"; Headers = @("Headline", "Description", "Image Src", "Image Alt", "Source Line"); Rows = $newsRows.ToArray() },
  @{ Name = "Sponsors Partners"; Headers = @("Grid", "Name", "Type Label", "Website", "Logo Src", "Logo Alt", "Source Line"); Rows = $sponsorRows.ToArray() },
  @{ Name = "Contact"; Headers = @("Field", "Label", "Value"); Rows = $contactRows.ToArray() },
  @{ Name = "Images Referenced"; Headers = @("Src", "Alt Text", "Kind", "Loading", "Source Line"); Rows = $imageRows.ToArray() },
  @{ Name = "Links Referenced"; Headers = @("Link Text", "Href", "Aria Label", "Source Line"); Rows = $linkRows.ToArray() },
  @{ Name = "Photo Library"; Headers = @("Path", "File Type", "Bytes", "Modified"); Rows = $photoRows.ToArray() }
)

$fullOutputPath = Join-Path (Split-Path $htmlFullPath) $OutputPath
if (Test-Path $fullOutputPath) { Remove-Item -LiteralPath $fullOutputPath -Force }

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::Open($fullOutputPath, [System.IO.Compression.ZipArchiveMode]::Create)
try {
  Add-ZipText $zip "_rels/.rels" '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

  $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
  $workbookSheets = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
  $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'

  for ($i = 0; $i -lt $sheets.Count; $i++) {
    $sheetId = $i + 1
    $sheet = $sheets[$i]
    $escapedName = [System.Security.SecurityElement]::Escape($sheet.Name)
    $contentTypes += "<Override PartName=""/xl/worksheets/sheet$sheetId.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>"
    $workbookSheets += "<sheet name=""$escapedName"" sheetId=""$sheetId"" r:id=""rId$sheetId""/>"
    $rels += "<Relationship Id=""rId$sheetId"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$sheetId.xml""/>"
    Add-SheetXml $zip "xl/worksheets/sheet$sheetId.xml" $sheet.Headers $sheet.Rows
  }

  $contentTypes += '</Types>'
  $workbookSheets += '</sheets></workbook>'
  $rels += "<Relationship Id=""rId$($sheets.Count + 1)"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/></Relationships>"

  Add-ZipText $zip "[Content_Types].xml" $contentTypes
  Add-ZipText $zip "xl/workbook.xml" $workbookSheets
  Add-ZipText $zip "xl/_rels/workbook.xml.rels" $rels
  Add-ZipText $zip "xl/styles.xml" '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'
}
finally {
  $zip.Dispose()
}

Write-Host "Created $fullOutputPath"
