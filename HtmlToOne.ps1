<#
    .NOTES
        AUTHORS : 
            Oleg_Astakhov@epam.com
        
        DATE: 11/1/2021
        Version: 3.0
        
        Please fill mandatory constants correctly.
        To allow execution of scripts type in console: Set-ExecutionPolicy Unrestricted
        
        (c)EPAM Systems Co.
        
    .SYNOPSIS
        Get release number, parse the page in Confluence about release and put found information into OneNote file.
        Using: HtmlToOne.ps1 -user <username> -pass <password> -release "Release+<release num>" -onefile <full path to onenote file>
        Example: HtmlToOne.ps1 - user Testuser -pass Testpassword -releae "Release+21.11.02"
        
    .DESCRIPTION
        Get HML-file, parse the file and put found information into OneNote file

    .EXAMPLE

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter (Mandatory=$true)]
    [string] $email = "",

    
    [Parameter (Mandatory=$true)]
    [string] $release = "",

    [Parameter (Mandatory=$true)]
    [string] $oneFile = "",

    [Parameter (Mandatory=$true)]
    [string] $templateFile = ""

)

$cred = $Credential
#Test-Path $onefile
#Test-Path $templateFile


$Headers = @{'Authorization' = "Basic "+[System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(($cred.UserName+":"+[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($cred.Password)) )))
            'X-Atlassian-Token' = 'nocheck'
        }

$confluenceUrl = "http://docs.bdainc.com/pages/viewpage.action?spaceKey=BC&title="+$release
$webContent = Invoke-WebRequest -Uri $confluenceUrl  -Headers $Headers -UseBasicParsing
$source = $webContent.Content

$html = New-Object -ComObject "HTMLFile";
try 
{
    $html.IHTMLDocument2_write($source);
}
catch 
{
    $encoded = [Text.Encoding]::Unicode.GetBytes($source)
    $html.write($encoded)
}



$UatAndProdDeploymentValue = ""
$EnviromentValue = ""
$FullBuildValue = ""
$DeployItemStockWsrvValue = ""
$DeployCCPaymentAzureValue = ""
$TaskEngineValue = ""
$ActualVersionValue = ""

$source = $source -replace "<span style=""color: rgb\(23,43,77\);"">", ""
$source = $source -replace "</span><span style=""color: rgb\(0,0,0\);"">", ""
$source = $source -replace "<span style=""color: rgb\(0,0,0\);"">", ""
$source = $source -replace "</span>|<span>", ""

$source = $source -replace "</th><td class=""confluenceTd""><div class=""content-wrapper""><p class=""confluence-link"">", ""

foreach($x in $source.split("<"))  
{
    if ($x.contains("FullBuild:") -eq $true) 
    { 
        $FullBuildValue = $x.split(">")[1].Split("/")[1]
    }

    if ($x.contains("FullBuild_Second:") -eq $true) 
    { 
        $FullBuildValue = $x.split(">")[1]
        $FullBuildValue = $FullBuildValue -replace "FullBuild_Second:\s+", ""
    }

    if ($x.contains("FB_BDAC") -eq $true) 
    { 
        $FullBuildValue = $x.split(">")[1].Split("/")[1]
    }

    if ($x.contains("DeployItemStockWsrv:") -eq $true) 
    { 
        $DeployItemStockWsrvValue = $x.split(">")[1]
        $DeployItemStockWsrvValue = $DeployItemStockWsrvValue -replace "DeployItemStockWsrv:\s+", ""
        $DeployItemStockWsrvValue = $DeployItemStockWsrvValue -replace "origin/", ""
    }

    if ($x.contains("deployCCPaymentAzure:") -eq $true) 
    { 
        $DeployCCPaymentAzureValue =  $x.split(">")[1]
        $DeployCCPaymentAzureValue = $DeployCCPaymentAzureValue -replace "deployCCPaymentAzure:", ""
        $DeployCCPaymentAzureValue = $DeployCCPaymentAzureValue -replace "origin/", ""
    }

    if ($x.contains("TaskEngine:") -eq $true) 
    { 
        $TaskEngineValue = $x.split(">")[1]
        $TaskEngineValue = $TaskEngineValue -replace "TaskEngine:", ""
    }

    if ($x.contains("Released on:") -eq $true) 
    { 
        $EnviromentValue = $x.split(">")[1]
        $EnviromentValue = $EnviromentValue -replace "Released on:", ""
        $EnviromentValue = $EnviromentValue.Split("at")[0]
    }
}

$ActualVersionValue = $FullBuildValue.Split("_")[0] + "_" + $FullBuildValue.Split("_")[1]
$UatAndProdDeploymentValue = $ActualVersionValue + " to UAT and PROD"



$OneNoteResult = New-Object -ComObject OneNote.Application

[ref]$xmlSection = ""
$OneNoteResult.OpenHierarchy($oneFile, "", $xmlSection, [Microsoft.Office.Interop.OneNote.CreateFileType]::cftSection)

[ref]$newPageID = ""
$OneNoteResult.CreateNewPage($xmlSection.Value, [ref]$newPageID, [Microsoft.Office.Interop.OneNote.NewPageStyle]::npsDefault)

[ref] $newPageXml = ""
$OneNoteResult.GetPageContent($newPageID.Value, [ref] $newPageXml, [Microsoft.Office.Interop.OneNote.PageInfo]::piAll)

$newPageContent = [xml]$newPageXml.Value



Write-Host "Full Build:"  $FullBuildValue
Write-Host "DeployItemStockWsrv:" $DeployItemStockWsrvValue
Write-Host "DeployCCPaymentAzure:" $DeployCCPaymentAzureValue
Write-Host "TaskEngine:" $TaskEngineValue
Write-Host "Enviroment:" $EnviromentValue
Write-Host "UAT and PROD:" $UatAndProdDeploymentValue
Write-Host "ActualVersion:" $ActualVersionValue


$template = [System.IO.File]::ReadAllText($templateFile)
$template = $template -replace "{{PAGEID}}", $newPageContent.Page.ID
$template = $template -replace "{{RELEASEUATANDPRD}}", $UatAndProdDeploymentValue
$template = $template -replace "{{FULLBUILDVALUE}}", $FullBuildValue
$template = $template -replace "{{ENVIRONMENTVALUE}}", $EnviromentValue
$template = $template -replace "{{STOCKWSRVVALUE}}", $DeployItemStockWsrvValue
$template = $template -replace "{{DEPLOYCCPAYMENTAZUREVALUE}}", $DeployCCPaymentAzureValue
$template = $template -replace "{{TASKENGINEVALUE}}", $TaskEngineValue
$template = $template -replace "{{ACTUALVERSIONVALUE}}", $ActualVersionValue

if ($release -ne "") {

    $releaseUrl = $release -Replace " ","_"
    $urlSql = "https://bitbucket.org/oastakhov/bdadev/src/" +  $releaseUrl + "/bdacommerce/Application/ClientData/DbScripts/" + $releaseUrl + ".sql"
    Write-Host $urlSql
    try{
        $a = Invoke-WebRequest -Uri $urlSql -UseBasicParsing 
        if ($a.StatusCode -eq 200) {
            Write-Host $urlSql
            $template = $template -replace "{{8}}", $urlSql           
        }
    }
    catch{
        $_.Exception.Response.StatusCode
    }
       
}
write-host "ready to save results"
try {
$OneNoteResult.UpdatePageContent($template.ToString()) 
}
catch {
    throw  $_.Exception.Response.StatusCode
    write-host "error writing to onenote"
}

Start-Sleep -s 10
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($OneNoteResult)
Start-Sleep -s 30
Send-MailMessage -Encoding UTF8 -From "notify@bdainc.com" -Subject "Roadmap for $release" `
-SmtpServer iwebsmtp.bdainc.bda -To $email `
 -Body 'OneNote file with release information' -Attachments $oneFile
 Start-Sleep -s 10

 if (Test-Path $oneFile) {
    Remove-Item $oneFile
  }
                                                                    
