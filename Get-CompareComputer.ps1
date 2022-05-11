<#
#### requires ps-version 5.1 ####
<#
.SYNOPSIS
This script compares software installed on one Windows computer to one or many computers.
.PARAMETER sourceComputer
This is the computer to compare all others against
.PARAMETER compareComputers
This is a comma separated string of one or more computers. 
.PARAMETER outputPath
This is the folder where the excel output file will be stored.
.INPUTS
None
.OUTPUTS
The output will be the excel spreadsheet containing the data. 
The first worksheet will be all the software installed on the source computer.
The next worksheets will show what is missing on the compare computers compared to the source computer.
    There will be two worksheets for each computer. One will omit msu packages and the other will show all.
The final worksheet combines all the compare computer results into one sheet.

Path to excel workbook is output to console as well.

.NOTES
   Version:        1.0
   Author:         Andy Orszula
   LinkedIn:       https://www.linkedin.com/in/orszula/
   Creation Date:  Wednesday, April 20th 2022, 11:23:48 am
   File: Get-CompareComputer.ps1
   Copyright (c) 2022 Kent Corporation

HISTORY:
Date      	          By	Comments
----------	          ---	----------------------------------------------------------
05/11/2022            AO    First version released to github   
.LINK
   
.COMPONENT
 Required Modules: 
    join-object
    importexcel

.LICENSE
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the Software), to deal
in the Software without restriction, including without limitation the rights
to use copy, modify, merge, publish, distribute sublicense and /or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED AS IS, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 
.EXAMPLE
PS> Get-CompareComputer -sourceComputer "foo01" -compareComputers foo02,foo03 -outputPath "c:\temp"
        c:\temp\Comparison-Output-05-11-2022_12_40.xlsx
PS> Get-CompareComputer -sourceComputer "foo01" -compareComputers foo02 -outputPath "c:\temp"
        c:\temp\Comparison-Output-05-11-2022_12_44.xlsx

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String]
    $sourceComputer,
    [Parameter(Mandatory)]
    [String]
    $compareComputers,
    [Parameter(Mandatory)]
    [String]
    $outputPath
    
)

#Check if required PowerShell modules are installed
$modules = @("Join-Object","ImportExcel")
$modules | ForEach-Object{
    
    try{

        Get-InstalledModule -Name $_ -ErrorAction Stop | out-null

    }
    catch{

        Write-Host "MODULE IS NOT INSTALLED. UNABLE TO CONTINUE" -ForegroundColor Yellow -BackgroundColor Blue 
        Write-Host $_ -ForegroundColor Yellow -BackgroundColor Blue 
        pause
        exit

    }

}

$computerNames = $compareComputers -split ","

#Excel OutPut Path  
    
if(!(test-path -Path $outputPath)){

    Write-Host "OUTPUT PATH DOES NOT EXIST. UNABLE TO CONTINUE" -ForegroundColor Yellow -BackgroundColor Blue 
    pause
    exit
}

$outputFile = "Comparison-Output-" + 
(get-date -f MM-dd-yyyy_HH_mm) +
 ".xlsx"
$outputPath = Join-Path -Path $outputPath -ChildPath $outputFile


function Get-RemotePackages{

    param(

    # Parameter help description
    [Parameter(Mandatory = $true)]
    [String]
    $computerName

    )

    $Packages = Invoke-Command -ComputerName $ComputerName -ScriptBlock {

        Get-Package
    
    
    }

    return $Packages
}

try{

    $packagesSourceObjects = Get-RemotePackages -ComputerName $sourceComputer -ErrorAction Stop

}
catch{

    Write-Host "PROBLEM CONNECTING TO SOURCE COMPUTER" -ForegroundColor Yellow -BackgroundColor Blue 
    Write-Host $_ -ForegroundColor Yellow -BackgroundColor Blue 
    pause
    exit

}


$packagesSourceObjects | Select-Object PSComputerName, CanonicalId  | 
Export-Excel -Path $outputPath -WorksheetName "SourceComputer" -TitleFillPattern LightGray -TableStyle Light1 -AutoSize

$allComputerData = @()
$computerNames | Sort-Object | ForEach-Object{

    $packagesCompareObject = $null
    $computerName = $_
    
    try{

        $packagesSourceObjects = $packagesCompareObject = Get-RemotePackages -ComputerName $computerName -ErrorAction Stop

        $comparison = Join-Object -Left $packagesSourceObjects `
        -Right $packagesCompareObject `
        -LeftJoinProperty CanonicalId,Installed `
        -RightJoinProperty CanonicalId,Installed `
        -Type AllInLeft `
        -LeftProperties PSComputerName,CanonicalId `
        -RightProperties PSComputerName,CanonicalId `
        -Prefix "Diffright-" `
        -ErrorAction Stop

        $sheet = $computername + "-all"
        $comparison | Select-Object PSComputerName, CanonicalId, CompareComputer | 
        Export-Excel -Path $outputPath -WorksheetName $sheet -TitleFillPattern LightGray -TableStyle Light1 -AutoSize
        
        $allComparisonFiltered = $comparison |
            Where-Object{
    
                $_.CanonicalId -notlike "msu:*"
    
            } | Where-Object{$null -eq $_."Compare-PSComputerName"}
        
                
        $updatesRemovedSheet = $computername + "-nomsu"
        $allComparisonFiltered | Select-Object PSComputerName, CanonicalId, CompareComputer | 
        Export-Excel -Path $outputPath -WorksheetName $updatesRemovedSheet -TitleFillPattern LightGray -TableStyle Light1 -AutoSize
        
        $comparison | ForEach-Object{
    
            $_ | Add-Member -NotePropertyName CompareComputer -NotePropertyValue $computerName
        
        }

        $allComputerData += $comparison
    
    }
    catch{
    
        Write-Host "PROBLEM CONNECTING TO COMPARE COMPUTER" -ForegroundColor Yellow -BackgroundColor Blue 
        Write-Host $_ -ForegroundColor Yellow -BackgroundColor Blue 
          
    }


}

$allComputerData | Select-Object PSComputerName, CanonicalId, CompareComputer | Sort-Object CompareComputer |
Export-Excel -Path $outputPath -WorksheetName "AllComputers" -TitleFillPattern LightGray -TableStyle Light1 -AutoSize

write-host $outputPath 
pause

