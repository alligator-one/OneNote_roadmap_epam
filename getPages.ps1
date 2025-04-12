Param(
    [Parameter (Mandatory=$false)]
    [string] $one = "",

    [Parameter (Mandatory=$false)]
    [string] $template = ""
)


$OneNote = New-Object -ComObject OneNote.Application

$OneNoteFilePath = $one
[ref]$oneNoteID = ""
[xml]$Hierarchy = ""
$OneNote.OpenHierarchy($OneNoteFilePath, "", $oneNoteID, [Microsoft.Office.Interop.OneNote.CreateFileType]::cftSection)
$OneNote.GetHierarchy($oneNoteID.Value, [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$Hierarchy)

$oneNoteID.Value
ForEach($page in $Hierarchy.Section.Page)
{
    [ref]$PageXML = ''
    $OneNote.GetPageContent($page.ID, [ref]$PageXML, [Microsoft.Office.Interop.OneNote.PageInfo]::piAll)
    
    [System.IO.File]::AppendAllText($template, $PageXML.Value)
    break
}

$templateBody = [System.IO.File]::ReadAllText($template)
$templateBody = $templateBody -replace """[{(]?[0-9A-F]*[-]?(?:[0-9A-F]*[-]?)*[0-9A-F]*[)}][{(]?[0-9A-F]*[-]?(?:[0-9A-F]*[-]?)*[0-9A-F]*[)}][{(]?[0-9A-F]*[-]?(?:[0-9A-F]*[-]?)*[0-9A-F]*[)}]""", ""
$templateBody = $templateBody -replace "objectID=", ""
$templateBody = $templateBody -replace " ID=", " ID=""{{PAGEID}}"""

[System.IO.File]::WriteAllText($template, $templateBody.ToString())