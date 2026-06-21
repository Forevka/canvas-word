# Generates scrape-sample.docx — a fixture for testing nested outer->inner content-control
# (w:sdt) round-trip fidelity in canvas-word.
#
# Structure:
#   - prose paragraph (unwrapped)
#   - OUTER block SDT (a "section", tagged like AppraiSys LazyContentDataModel {Name,Hash})
#       - paragraph with an INNER inline field-binding SDT (target=appraiseRequest, arKey=FileNumber)
#       - 2-col table whose VALUE cells each hold an INNER inline field-binding SDT:
#           * Appraisal Fee  -> target=formFieldValue (orig raw "200", display "$200.00")
#           * Owner Name     -> target=formSubmission  (orig "John Doe")
#           * Property City  -> target=appraiseRequest (arKey=PropertyLocation.City)
#   - NON-NESTED inline field-binding SDT (for comparison)
#
# Run:  pwsh test-fixtures/scrape/make-fixture.ps1

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$out  = Join-Path $here 'scrape-sample.docx'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# --- helpers ---------------------------------------------------------------
function Xml-Escape([string]$s) {
  $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;')
}

# Build a FieldBinding JSON object and return it as an XML-attribute-escaped string.
function Fb-Tag([hashtable]$o) {
  # stable key order for readability
  $order = @('k','target','type','orig','storyId','fieldName','arKey','itemKey')
  $parts = @()
  foreach ($key in $order) {
    if ($o.ContainsKey($key)) {
      $v = $o[$key]
      if ($null -eq $v)            { $parts += "`"$key`":null" }
      elseif ($v -is [int])        { $parts += "`"$key`":$v" }
      else                         { $parts += "`"$key`":`"$v`"" }
    }
  }
  $json = '{' + ($parts -join ',') + '}'
  Xml-Escape $json
}

# An inline field-binding content control wrapping a single value run.
function Inner-Sdt([string]$alias, [hashtable]$binding, [string]$display) {
  $tag = Fb-Tag $binding
  $al  = Xml-Escape $alias
  $tx  = Xml-Escape $display
@"
<w:sdt><w:sdtPr><w:alias w:val="$al"/><w:tag w:val="$tag"/></w:sdtPr><w:sdtContent><w:r><w:t xml:space="preserve">$tx</w:t></w:r></w:sdtContent></w:sdt>
"@
}

# A label/value table row; the value cell embeds an inner field-binding SDT.
function Table-Row([string]$label, [string]$innerSdtXml) {
  $lab = Xml-Escape $label
@"
<w:tr>
  <w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">$lab</w:t></w:r></w:p></w:tc>
  <w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p>$innerSdtXml</w:p></w:tc>
</w:tr>
"@
}

# --- bindings --------------------------------------------------------------
$fbFileNumber = @{ k='fb'; target='appraiseRequest'; type='text'; orig='CM099026-2876-6323'; arKey='FileNumber' }
$fbFee        = @{ k='fb'; target='formFieldValue'; type='currency'; orig='200'; storyId=42; fieldName='executive_summary_appraisal_fee' }
$fbOwner      = @{ k='fb'; target='formSubmission'; type='text'; orig='John Doe'; storyId=42; fieldName='executive_summary_owner' }
$fbCity       = @{ k='fb'; target='appraiseRequest'; type='text'; orig='Panama City Beach'; arKey='PropertyLocation.City' }
$fbStandalone = @{ k='fb'; target='formFieldValue'; type='number'; orig='999'; storyId=7; fieldName='standalone_test_value' }

$innerFileNumber = Inner-Sdt 'File Number' $fbFileNumber 'CM099026-2876-6323'
$innerFee        = Inner-Sdt 'Appraisal Fee' $fbFee '$200.00'
$innerOwner      = Inner-Sdt 'Owner Name' $fbOwner 'John Doe'
$innerCity       = Inner-Sdt 'Property City' $fbCity 'Panama City Beach'
$innerStandalone = Inner-Sdt 'Standalone Value' $fbStandalone '999'

$rows = (Table-Row 'Appraisal Fee' $innerFee) +
        (Table-Row 'Owner Name'    $innerOwner) +
        (Table-Row 'Property City' $innerCity)

# Outer-section tag mirrors AppraiSys LazyContentDataModel { Name, Hash }
$outerTagJson = Xml-Escape '{"Name":"<subject_property_information_substitutor>","Hash":"abc123"}'

# --- document.xml ----------------------------------------------------------
$documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve">Scrape Fixture — Subject Property Information</w:t></w:r></w:p>

    <w:sdt>
      <w:sdtPr>
        <w:alias w:val="Subject Property Information (section)"/>
        <w:tag w:val="$outerTagJson"/>
      </w:sdtPr>
      <w:sdtContent>
        <w:p><w:r><w:t xml:space="preserve">File number: </w:t></w:r>$innerFileNumber</w:p>
        <w:tbl>
          <w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          </w:tblBorders></w:tblPr>
          <w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
          $rows
        </w:tbl>
      </w:sdtContent>
    </w:sdt>

    <w:p><w:r><w:t xml:space="preserve">Standalone (not nested): </w:t></w:r>$innerStandalone</w:p>

    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>
"@

$contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"@

$rels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

$docRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@

$styles = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>
"@

# --- zip into .docx --------------------------------------------------------
if (Test-Path $out) { Remove-Item $out -Force }
$fs = [System.IO.File]::Open($out, [System.IO.FileMode]::CreateNew)
try {
  $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    function Add-Entry($zip, $name, $content) {
      $entry  = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), [System.Text.UTF8Encoding]::new($false))
      try { $writer.Write($content) } finally { $writer.Dispose() }
    }
    Add-Entry $zip '[Content_Types].xml'        $contentTypes
    Add-Entry $zip '_rels/.rels'                $rels
    Add-Entry $zip 'word/document.xml'          $documentXml
    Add-Entry $zip 'word/_rels/document.xml.rels' $docRels
    Add-Entry $zip 'word/styles.xml'            $styles
  } finally { $zip.Dispose() }
} finally { $fs.Dispose() }

Write-Host "Wrote $out"
