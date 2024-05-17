# Functions==========================================================================================
function getHierarchyItem($hierarchy) {
  $data = @()

  foreach ($items in $hierarchy.ChildNodes) {
    $data += getHierarchyChildItem $items
  }
  return $data
}

function getHierarchyChildItem($hierarchyItems) {

  $data = @()

  foreach ($item in $hierarchyItems) {
    $data_item = New-Object PSObject | Select-Object name, ID

    if ($item.name -eq 'Deleted Pages') {
      continue;
    }
    $data_item.name = $item.name
    $data_item.ID = $item.ID
    $data += $data_item
  
    if ($item.HasChildNodes) {
      getHierarchyItem $item
    }
  }
  return $data
}

function isStartedProcess() {
  $proc = Get-Process "ONENOTE" -ErrorAction SilentlyContinue
  
  if ($proc) {
    return $true
  }
  else {
    return $false
  }
}

function disposeComObject($OneNote) {
  # Dispose COM object
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($OneNote) | Out-Null
}

# Main process==========================================================================================

$ConfigPath = "./config.ini"
$Param = @{}
Get-Content $ConfigPath | % { $Param += ConvertFrom-StringData $_ }

$SectionName = $PARAM.SECTION_NAME

# Initialize
$OneNote = New-Object -ComObject OneNote.Application
Add-Type -assembly Microsoft.Office.Interop.OneNote
Add-Type -AssemblyName System.Xml.Linq

  
$OneNoteID = ""
[xml]$Hierarchy = ""
  
# Get OneNote hierarcy structure
$OneNote.GetHierarchy(
  $OneNoteID,
  [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages,
  [ref]$Hierarchy,
  [Microsoft.Office.InterOp.OneNote.XMLSchema]::xsCurrent)
$HierarchyItem = getHierarchyItem $Hierarchy
    
# Get disignated section id
[string]$SectionID = ($HierarchyItem | Where-Object { $_.name -eq $SectionName }).ID
    
if ($null -eq $SectionID) {
      
  Write-Error 'Section ID is not found'
  disposeComObject $OneNote
  Pause
  return
}
    
# Check today's page has been already created
$date = Get-Date -Format "yyyy-MM-dd"
$TodaysPage = ($HierarchyItem | Where-Object { $_.name -eq $date })
    
# Today's page wad already created
if (($null -ne $TodaysPage)) {
  Write-Error "Today's page wad already created."
  disposeComObject $OneNote
  Pause
  return
}
    
# Create page to disignated section
$pbstrPageID = ""
$OneNote.CreateNewPage(
  $SectionID.Trim(),
  [ref]$pbstrPageID,
  [Microsoft.Office.Interop.OneNote.NewPageStyle]::npsBlankPageWithTitle
)

# Get new page xml
[ref]$NewPageXML = ''
$OneNote.GetPageContent($pbstrPageID, [ref]$NewPageXML, [Microsoft.Office.Interop.OneNote.PageInfo]::piAll, [Microsoft.Office.InterOp.OneNote.XMLSchema]::xsCurrent)
$xDoc = [System.Xml.Linq.XDocument]::Parse($NewPageXML.Value)

# Get title element
$title = $xDoc.Descendants() | Where-Object -Property Name -Like -Value '*}T'
if (-not $title) {
  Write-Error 'Error: can not find title element'
  disposeComObject $OneNote
  Pause
  return
}

# Set page title
$title.Value = "$date"

# Update
$onenote.UpdatePageContent($xDoc.ToString())

# Dispose COM object
disposeComObject $OneNote

Start-Sleep -s 1

$isStarted = isStartedProcess
if ($isStarted) {
  return
}

# Open OneNote
Start-Process -FilePath "ONENOTE"