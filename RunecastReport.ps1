<#
.SYNOPSIS
  Runecast Report
.DESCRIPTION
  Generate a report in an Excel file from Runecast using API
.PARAMETER Server
    Runecast server
.PARAMETER Token
    Runecast Authorization Token
.PARAMETER Output
    Excel filename
.INPUTS
  none
.OUTPUTS
  none
.NOTES
  Version:        2.0
  Author:         Julien Mousqueton
  Creation Date:  01/06/2020
  Original:       vMan.ch
  Purpose/Change: Initial script development
  Requirement:    Make sure to install ImportExcel module
                  Install-Module ImportExcel

.EXAMPLE
  .\RunecastReposrt.ps1 -Server runecast.local -Token '1234567ab-89bc-012d-3e45-678f9gh12345' -Output 'Extract.xlsx'
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
# None #
#----------------------------------------------------------[Declarations]----------------------------------------------------------
# Script Version
$ScriptVersion = "2.0"
#-----------------------------------------------------------[Functions]------------------------------------------------------------
# None #
#-----------------------------------------------------------[Execution]------------------------------------------------------------

param
(
    [String]$Server,
    [String]$Token,
    [String]$Output
)

#Stuff for Invoke-RestMethod
$ContentType = "application/json"
$header = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$header.Add("Accept", 'application/json')
$header.Add("Authorization", $Token)
$header.Add("User-Agent", 'RunecastExtractor/2.0')


#Get a list of every Issue


    $IssueList = @()

    # I've got an internal error when I used only  /rc2/api/v1/issues
    # I've add ?type=KB and everything goes smoothly
    $IssueUrl = 'https://'+$Server+'/rc2/api/v1/issues?type=KB'

    $issues = Invoke-RestMethod -Uri $IssueUrl -Method Get -Headers $header -ContentType $ContentType -SkipCertificateCheck

    ForEach ($issue in $issues.issues){

        $IssueList += New-Object PSObject -Property @{

                    id = $issue.id
                    affects = $issue.affects
                    appliesTo = $issue.appliesTo
                    severity = $issue.severity
                    type = $issue.type
                    title = $issue.title
                    url = $issue.url
                    annotation = $issue.annotation
                    updatedDate = $issue.updatedDate
                    stigid = $issue.stigid
                    vulnid = $issue.vulnid
                    checkDescription = $issue.checkDescription
                    fixDescription = $issue.fixDescription
                    stigSection = $issue.stigSection

        }

    }


#Get a list of VC's

    $VCList = @()

    $VCUrl = 'https://'+$Server+'/rc2/api/v1/vcenters'

    $VCs = Invoke-RestMethod -Uri $VCUrl -Method Get -Headers $header -ContentType $ContentType -SkipCertificateCheck

    ForEach ($VC in $VCs.vcenters){

        $VCList += New-Object PSObject -Property @{

                    vcUid = $VC.uid
                    address = $VC.address
        }

    }


#Get a list of results

    $ResultsList = @()

    $ResultsUrl = 'https://'+$Server+'/rc2/api/v1/results'

    $results = Invoke-RestMethod -Uri $ResultsUrl -Method Get -Headers $header -ContentType $ContentType -SkipCertificateCheck

    ForEach ($Result in $Results.Results.issues){

        $id = $Result.id
        $status = $Result.Status

        ForEach ($affectedObject in $Result.affectedObjects){

            $ResultsList += New-Object PSObject -Property @{

                    id = $id
                    Name = $affectedObject.Name
                    vcUid = $affectedObject.vcUid
                    moid = $affectedObject.moid
        }
       }
    }

#Export it all to Excel

$IssueList | Select id,affects,appliesTo,severity,type,title,url,annotation,updatedDate,stigid,vulnid,stigSection | export-excel $Output -WorkSheetname Issues

$VCList | Select vcUid,address | export-excel $Output -WorkSheetname vCenters

$ResultsList | Select id,Name,vcUid,moid | export-excel $Output -WorkSheetname Results
