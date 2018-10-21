# Office 365 Message Center Digest

Get-MessageCenterDigest.ps1 is a PowerShell script that provides an email and HTML report of the messages in the Message Center of an Office 365 tenant.

The script will store information about the Message Center messages for a tenant in a file named *MessageCenterArchive-tenantdomain.xml*, located in the same folder as the script. The first time you run the script, all messages will be reported as "New". On subsequent runs, the script will use the previous results to determine which messages are new or changed since the last time the script was run.

## Installation

This script has a dependency on the following functions and modules:

- [New-StoredCredential & Get-StoredCredential](http://practical365.com/blog/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks/)
- [O365ServiceCommunications module](https://github.com/mattmcnabb/O365ServiceCommunications)

To install the script:

1. Add the New-StoredCredential & Get-StoredCredential functions to your [PowerShell profile](http://practical365.com/exchange-server/create-powershell-profile/)
2. Install the O365ServiceCommunications module:

```
PS> Install-Module -Name O365ServiceCommunications
```

3. Download the latest release from the [TechNet Script Gallery](https://gallery.technet.microsoft.com/office/Office-365-Message-Center-de1f7e5a).
4. Unzip the files to a folder on the workstation or server where you want to run the script.
5. Rename *Get-MessageCenterDigest.xml.sample* to *Get-MessageCenterDigest.xml*
6. Edit *Get-MessageCenterDigest.xml* with appropriate email settings for your environment. If you exclude the SMTP server, the script will send the report email to the first MX record for the domain of the *To* address.
7. Create a new stored credential by running New-StoredCredential
8. Run the script using the usage examples below.

## Usage  

Run the script in a PowerShell console.

```
.\Get-MessageCenterDigest.ps1 -UseCredential admin@tenantname.onmicrosoft.com
```

Run the script with verbose output.

```
.\Get-MessageCenterDigest.ps1 -UseCredential admin@tenantname.onmicrosoft.com -Verbose
```

Run the script with manual SMTP settings that override the Get-MessageCenterDigest.xml configuration.

```
.\Get-MessageCenterDigest.ps1 -MailFrom reports@contoso.com -MailTo you@contoso.com -MailSubject "Your custom subject" -SmtpServer mail.contoso.com
```

Run the script, suppressing the email report, and generating a HTML file instead.

```
.\Get-MessageCenterDigest.ps1 -UseCredential admin@tenantname.onmicrosoft.com -NoEmail
```

## Credits

Written by: Paul Cunningham

Find me on:

* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	https://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

Check out my [books](https://paulcunningham.me/books/) and [courses](https://paulcunningham.me/training/) to learn more about Office 365 and Exchange Server.

Additional contributions:

- Ryan Mitchell (improved report styling)
