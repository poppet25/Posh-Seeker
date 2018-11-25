using namespace System.IO
using namespace System.Xml
using namespace DocumentFormat.OpenXml
using namespace DocumentFormat.OpenXml.Packaging
using namespace DocumentFormat.OpenXml.Spreadsheet

try{
    Add-Type -Path ./lib/Tesseract.dll
    Add-Type -Path ./lib/itext.kernel.dll
    Add-Type -Path ./lib/DocumentFormat.OpenXml.dll
    Import-Module -Name PSLiteDB
}catch{
    Write-Host $_
}

[String]$ssn = @(
    '\.*\d{3}-\d{2}-\d{4}\.*',
    '\.*(?!000)(?!666)(?<SSN3>[0-6]\d{2}|7(?:[0-6]\d|7[012]))([- ]?)(?!00)(?<SSN2>\d\d)\1(?!0000)(?<SSN4>\d{4})\.*',
    '\.*((?!000)(?!666)([0-6]\d{2}|7[0-2][0-9]|73[0-3]|7[5-6][0-9]|77[0-1]))(\s|\-)((?!00)\d{2})(\s|\-)((?!0000)\d{4})\.*',
    '\b(?!000)(?!666)(?!9)[0-9]{3}[ -]?(?!00)[0-9]{2}[ -]?(?!0000)[0-9]{4}\b'
) -join '|'

function SearchDOCX
{    
    Param(
        # Parameter help description
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]
        $File
    )
    
    [string]$StartPartRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    
    try{
        [Package]$myPackage = [Package]::Open($File, [FileMode]::Open, [FileAccess]::ReadWrite)
    
        foreach ($relationship in $myPackage.GetRelationshipsByType($StartPartRelationship))
        {
            [Uri]$documentUri = [PackUriHelper]::ResolvePartUri([Uri]::new("/", [UriKind]::Relative),$relationship.TargetUri)
            
            [PackagePart] $StartPart = $myPackage.GetPart($documentUri)
            [XmlReaderSettings] $ReaderSettings = [XmlReaderSettings]::new()
            $ReaderSettings.IgnoreComments = $true
            $ReaderSettings.IgnoreWhitespace = $true
            
            [XmlReader] $MyTextReader = [XmlReader]::Create($StartPart.GetStream(), $ReaderSettings)
            
                while ($MyTextReader.ReadToFollowing("w:t") -and ($MyTextReader.EOF -eq $false))
                {
                    $value = $MyTextReader.ReadString()

                    if ($value -match $regex)
                    {
                        $MyTextReader.Close()
                        return @{HasSSN=$true; Value=$value}
                    }
                }

                break
        }

        $myPackage.Close()
        return $false
    }
    catch{
        Write-Host $Error
    }
}

function SearchXLSX
{    
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]
        $File
    )
    try{
        [SpreadsheetDocument] $spreadsheetDocument = [SpreadsheetDocument]::Open($File, $false)
        [WorkbookPart]$workbookPart = $spreadsheetDocument.WorkbookPart
        [OpenXmlReader] $reader = [OpenXmlReader]::Create($workbookPart.SharedStringTablePart)

        while ($reader.Read())
        {
            if ($reader.ElementType -eq [SharedStringItem])
            {
                [SharedStringItem]$ssi = $reader.LoadCurrentElement()
                New-Object psobject -Property @{ Value = $ssi.Text.Text }
            }
        }

        $spreadsheetDocument.Close()
    }
    catch{

    }
}

$directories = [Directory]::EnumerateDirectories("/Users",[SearchOption]::AllDirectories)

$directories | %{ [Directory]::GetFiles($_) } | ?{ $_.Name.EndsWith("DOCX")} | SearchXLSX -SearchText "ssn" -File "$PSScriptRoot/docs/Book1.xlsx"