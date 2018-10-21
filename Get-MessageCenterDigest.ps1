<#
.SYNOPSIS
Get-MessageCenterDigest.ps1 - Generates a report of Office 365 Message Center messages.

.DESCRIPTION 
This script provides an email and HTML report of the messages in the Message Center of an Office 365 tenant.

.OUTPUTS
Email to defined recipient(s).

.PARAMETER UseCredential
Credentials to pass to New-SCSession. Requires that the Get-StoredCredential
function (http://practical365.com/blog/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks/)
be available.

.EXAMPLE
.\Get-MessageCenterDigest.ps1 -UseCredential admin@tenantname.onmicrosoft.com

.EXAMPLE
.\Get-MessageCenterDigest.ps1 -Verbose

.LINK
https://github.com/cunninghamp/Office365MessageCenterDigest

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	https://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

Additional contributions:

- Ryan Mitchell (improved report styling)

Version history:
V1.00, 25/01/2017 - Initial version
V1.01, 09/02/2017 - Updated with better report styling


License:

The MIT License (MIT)

Copyright (c) 2016 Paul Cunningham

Permission is hereby granted, free of charge, to any person obtaining a copy 
of this software and associated documentation files (the "Software"), to deal 
in the Software without restriction, including without limitation the rights 
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
copies of the Software, and to permit persons to whom the Software is 
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all 
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
DEALINGS IN THE SOFTWARE.

#>

#requires -Modules O365ServiceCommunications

[CmdletBinding()]
param (
        [Parameter(Mandatory=$true)]
        [string]$UseCredential,
        [Parameter(Mandatory=$false)]
        [string]$MailTo,
        [Parameter(Mandatory=$false)]
        [string]$MailFrom,
        [Parameter(Mandatory=$false)]
        [string]$MailSubject,
        [Parameter(Mandatory=$false)]
        [switch]$NoEmail

)


#region Variables


$now = Get-Date

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$NewMessages = @()
$ChangedMessages = @()
$UnchangedMessages = @()
$NewResults = @()
$LastResults = @()
$NewMessageCount = 0
$ChangedMessageCount = 0
$UnchangedMessageCount = 0

$tenant = $usecredential.Split("@")[1]

$XMLFileName = "$($myDir)\MessageCenterArchive-$($tenant).xml"
$HtmlReportFileName = "$($myDir)\MessageCenterDigest-$($tenant).html"



#endregion Variables


#region Script Config

# Import settings from configuration file
$ScriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name))
$ConfigFile = Join-Path -Path $MyDir -ChildPath "$ScriptName.xml"
if (-not (Test-Path $ConfigFile)) {
    # Config file not found, make sure the minimum mandatory variables have been provided
    if (-not ($MailTo -and $MailFrom)) {
        # Config file not found, and no required parameters were provided 
        throw "Could not find configuration file! Be sure to rename the '$ScriptName.xml.sample' file to '$ScriptName.xml'. Alternatively, provide -Mail* arguments (see Get-Help for more information)." 
    } else {
        # Config file not found, but parameters were provided

    }
} else {
    $settings = ([xml](Get-Content $ConfigFile)).Settings
}

# If the $smtpSettings.SmtpServer value is either "" or $null, the script
# will attempt to automatically derive the SMTP server from the recipient
# domain's MX records
# If values are provided as arguments, use those, otherwise retrieve from XML file
$smtpsettings = @{
	To =  if ($MailTo) { $MailTo } else { $settings.EmailSettings.To }
	From = if ($MailFrom) { $MailFrom } else { $settings.EmailSettings.From }
	Subject = "$(if ($MailSubject) { $MailSubject } elseif ($settings.EmailSettings.Subject) { $settings.EmailSettings.Subject } else { "Office 365 Message Center Digest" } ) - $now - $tenant"
	SmtpServer = if ($SmtpServer) { $SmtpServer } else { $settings.EmailSettings.SmtpServer }
	}


# If there's no SMTP Server specified, attempt to derive one from MX records
if ([string]::IsNullOrWhiteSpace($settings.EmailSettings.SmtpServer)) {
    Write-Verbose "No SMTP server was specified - deriving one from DNS"
    try {
        $recipientSmtpDomain = $smtpSettings.To.Split("@")[1]
        $MX = Resolve-DnsName -Name $recipientSmtpDomain -Type MX | 
            Where-Object {$_.Type -eq "MX"} | 
            Sort-Object Preference | 
            Select-Object -First 1 -ExpandProperty NameExchange
        Write-Verbose "Found MX record: '$MX'"
        $smtpsettings.SmtpServer = $MX
    } catch {
        throw "Unable to resolve SMTP Server and none was specified.`n$($_.Exception.Message)"
    }
}

#endregion Script Config


#region Functions

#This function generates the HTML for each message being added to a table
Function Get-MessageHtml()
{
  param( $message )
	
  $Messagehtml = $null

  $f = $Message.Messages -replace("`n",'<br>') -replace([char]8217,"'") -replace([char]8220,'"') -replace([char]8221,'"') -replace('\[','<b><i>') -replace('\]','</i></b>')

  $t = $Message.Title -replace("`n",'<br>') -replace([char]8217,"'") -replace([char]8220,'"') -replace([char]8221,'"')
  
  $u = switch ($Message."Urgency"){
    'Critical' {' style="color:#ffffff;background-color:#ff0000;font-weight:bold"'} #red backgound/white text/bold
    'High'     {' style="color:#ff0000;font-weight:bold"'}                          #red text/bold
    'Normal'   {' style="color:#000000"'}                                           #black text
    Default    {' style="color:#000080;font-weight:bold"'}                          #navy text/bold
  }

  $Messagehtml += "<tr>
                    <th colspan=""6"" style=""color:#ffffff;background-color:#000099""><H4>$($Message."Message ID") - $t</H4></th>
                    </tr>
                    <tr>
                    <th>Category</th>
                    <th>Published</th>
                    <th>Expires</th>
                    <th>Urgency</th>
                    <th>Action Type</th>
                    <th>Required By</th>
                    </tr>
                    <tr>
                    <td align=""center"">$($Message."Category")</td>
                    <td align=""center"">$($Message."Start Time")</td>
                    <td align=""center"">$($Message."End Time")</td>
                    <td align=""center""$u>$($Message."Urgency")</td>
                    <td align=""center"">$($Message."Action")</td>
                    <td align=""center"" style=""color:#ff0000""><b>$($Message."Action Required By")</b></td>
                    </tr>
                    <tr>
                    <td colspan=""6"">$f<br/><a href=""$($Message.Link)"">More Info</a></td>
                  </tr>
                  <tr><td colspan=""6"" style=""border:none""></td></tr>"
	
  return $Messagehtml
}

#endregion Functions


#region Main Script


#Check dependency

if (-not (Get-Command Get-StoredCredential -ErrorAction SilentlyContinue)) {
    throw "Add the PSStoredCredentials functions to your profile (Refer to: http://practical365.com/blog/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks/)"
}

# $isFirstRun is used to flag whether this is the first time the script has run
$isFirstRun = $false

#Check for previous results
if (Test-Path $XMLFileName) {
    
    #XML file found, ingest as last results
    $LastResults = Import-Clixml -Path $XMLFileName
    Write-Verbose "There were $($LastResults.Count) previous messages checked."
}
else {
    Write-Verbose "No previous results found."
    $isFirstRun = $true
}



# Check Message Center

$mycred = Get-StoredCredential -UserName $UseCredential

$MySession = New-SCSession -Credential $MyCred

$events = Get-SCEvent -EventTypes Message -SCSession $MySession

$LastRunIds = $LastResults."Message Id"

$EventIds = $Events.Id

#TODO - need to check if LastUpdatedTime has changed
#TODO - need to also include ActionRequiredByDate info
#TODO - MessageText is not well formatted in final report
foreach ($EventId in $EventIds) {
    Write-Verbose "Checking if $EventId has been seen before"

    $EventDetails = $Events | Where {$_.Id -eq $EventId}

    #Properties for a custom object to store details for the report and any changed properties
    $objectProperties = [ordered]@{
                                "Message ID" = $EventDetails.Id
                                "Title" = $EventDetails.Title
                                "Urgency" = $EventDetails.UrgencyLevel
                                "Start Time" = $EventDetails.StartTime
                                "End Time" = $EventDetails.EndTime
                                "Last Updated" = $EventDetails.LastUpdatedTime
                                "Category" = $EventDetails.Category
                                "Action" = $EventDetails.ActionType
                                "Action Required By" = $EventDetails.ActionRequiredByDate
                                "Messages" = $EventDetails.Messages.MessageText | Out-String
                                "Link" = $EventDetails.ExternalLink
				                }

    $messageObj = New-Object -TypeName PSObject -Property $objectProperties

    if (-not ($LastRunIds -icontains $EventId)) {
    
        Write-Verbose "Message ID $EventId is new"
        
        $NewMessages += $messageObj

        $NewMessageCount++
    }
    elseif ($LastRunIds -icontains $EventId) {
        
        $hasChanged = $false

        $LastEventDetails = $LastResults | Where {$_."Message ID" -eq $EventId}

        foreach ($Property in $ObjectProperties.Keys) {
            if ($($LastEventDetails.$Property) -ne $($messageObj.$Property)) {

                Write-Verbose "$Property has changed for $($EventDetails.Id)"
                $hasChanged = $true
            }
        }

        if ($hasChanged) {

            $ChangedMessages += $messageObj

            $ChangedMessageCount++
        }

    }
}

$NewMessages = $NewMessages | Sort "Last Updated"
$ChangedMessages = $ChangedMessages | Sort "Last Updated"

if ($($LastResults.Count) -gt 0) {
   $NewResults = $LastResults | Where {($NewMessages."Message ID" -notcontains $_."Message ID") -and ($ChangedMessages."Message ID" -notcontains $_."Message ID")}
   $UnchangedMessageCount = $NewResults.Count
}

foreach ($NewMessage in $NewMessages) {
    $NewResults += $NewMessage
}

foreach ($ChangedMessage in $ChangedMessages) {
    $NewResults += $ChangedMessage
}

Write-Verbose "There were $NewMessageCount new messages."
Write-Verbose "There were $ChangedMessageCount changed messages."
Write-Verbose "There were $UnchangedMessageCount unchanged messages."


Write-Verbose "Writing current events to XML file."
$NewResults | Export-CliXml $XMLFileName -Force


#region Email Report


#TODO - these styles are copied from another script so may need cleaning up
#HTML HEAD with styles
$htmlhead="<html>
			<style>
			BODY{font-family: Arial; font-size: 10pt;}
			H1{font-size: 22px;}
			H2{font-size: 18px; padding-top: 10px;}
			H3{font-size: 16px; padding-top: 8px;}
            H4{font-size: 12px; padding-top: 4px;}
			TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt; table-layout: fixed; width: 800px;}
            TABLE.summary{text-align: center; width: auto;}
			TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
			TD{border: 1px solid black; padding: 5px; vertical-align: top; }
			td.pass{background: #7FFF00;}
			td.warn{background: #FFE600;}
			td.fail{background: #FF0000; color: #ffffff;}
			td.info{background: #85D4FF;}
            ul{list-style: inside; padding-left: 0px;}
            .firstrunnotice { font-size: 14px; color: #4286f4; }
			</style>
			<body>"

#HTML TAIL
$htmltail = "</body>
			</html>"

$htmlIntro = "<h1>Office 365 Message Center Digest</h1>"

#New messages table
#$HtmlBody = $NewMessages | ConvertTo-Html -Fragment
if ($NewMessageCount -gt 0) {

    $NewMessagesTable = "<h3>New Messages</h3>
						<table>"

    foreach ($Message in $NewMessages) {

        $NewMessagesTable += (Get-MessageHtml $message)

    }

    $NewMessagesTable += "</table>"
}

#Changed messages table
if ($ChangedMessageCount -gt 0) {

    $ChangedMessagesTable = "<h3>Changed Messages</h3>
						<table>"

    foreach ($Message in $ChangedMessages) {

        $ChangedMessagesTable += (Get-MessageHtml $message)

    }

    $ChangedMessagesTable += "</table>"
}

$htmlemail = $htmlhead + $HtmlIntro + $NewMessagesTable + $ChangedMessagesTable + $htmltail


if (-not ($noEmail) -and (($NewMessageCount -gt 0) -or ($ChangedMessageCount -gt 0))) {
    try {
        Send-MailMessage @smtpsettings -Body $htmlemail -BodyAsHtml -ErrorAction STOP
        Write-Verbose "Email report sent."
    }
    catch {
        Write-Verbose "Email report not sent."
        throw $_.Exception.Message
    }
}

if (($noEmail) -and (($NewMessageCount -gt 0) -or ($ChangedMessageCount -gt 0))) {
    $htmlEmail | Out-File $HtmlReportFileName -Force
}

#endregion Email Report

#endregion Main Script
