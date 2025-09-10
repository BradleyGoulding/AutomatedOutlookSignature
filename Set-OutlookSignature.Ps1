[CmdletBinding()]
param (
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]
    $CompanyName,
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]
    $Website,
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("ascii", "utf8", "unicode", "utf32")]
    [string]
    $Encoding = "unicode",
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]
    $SignatureName = "Signature",
    [Parameter()]
    [string]
    $UseWordTemplate = "signature-template.docx"
)

<#
.SYNOPSIS
    Automated Outlook Signature

.DESCRIPTION
    Generates Outlook email signatures from user Active Directory information. 
    Supports both traditional HTML/TXT templates and Word document templates.
    
    The script can operate in two modes:
    1. Standard mode: Uses template.html and template.txt files
    2. Word Template mode: Uses signature-template.docx with Word automation

.PARAMETER CompanyName
    Override the company name from Active Directory

.PARAMETER Website
    Override the website URL from Active Directory

.PARAMETER Encoding
    Text encoding for output files (ascii, utf8, unicode, utf32)

.PARAMETER SignatureName
    Name for the signature files and Outlook signature setting (default: "Signature")

.PARAMETER UseWordTemplate
    Use Word document template instead of HTML/TXT templates. 
    Specify the template filename (e.g., "AEE Primary.docx") or leave empty to use "template.docx"
    Requires Microsoft Word to be installed

.EXAMPLE
    .\Set-OutlookSignature.Ps1
    Generate signatures using HTML/TXT templates

.EXAMPLE
    .\Set-OutlookSignature.Ps1 -UseWordTemplate "" -Verbose
    Generate signatures using default Word template (template.docx) with verbose output

.EXAMPLE
    .\Set-OutlookSignature.Ps1 -UseWordTemplate "AEE Primary.docx"
    Generate signatures using specific Word template file

.EXAMPLE
    .\Set-OutlookSignature.Ps1 -CompanyName "Acme Corp" -Website "www.acme.com"
    Override company name and website from AD values

.EXAMPLE
    .\Set-OutlookSignature.Ps1 -SignatureName "Corporate" -UseWordTemplate "AEE Primary.docx"
    Generate signatures using custom Word template with custom signature name

.NOTES
    Author           : CaptainQwerty
    Release Date     : 19/05/2024
    Script Version   : 6.0.0
    GitHub Repo      : https://github.com/CaptainQwerty/AutomatedOutlookSignature
    ReadMe           : https://github.com/CaptainQwerty/AutomatedOutlookSignature/blob/main/README.md
    
    Word Template Requirements:
    - Microsoft Word must be installed
    - Template file must exist in script directory (default: template.docx)
    - Template should use {{FieldName}} placeholders for dynamic content
#>

<#-----[ Latest Patch Notes ]-----#

Version 6.0.0
    * Added Word template support with -UseWordTemplate parameter
    * Word automation generates RTF, HTML, and TXT formats automatically
    * Maintains backward compatibility with HTML/TXT template method
    * Enhanced error handling and COM object cleanup
    * Professional signature formatting through Word's rich formatting
    * Template placeholder system: {{FieldName}} for dynamic content
    * Automatic address formatting and empty field handling

Version 5.0.0
    * New script layout utilising functions to help with readability
    * Now Supports Verbose
    * Support for paramaters with the script execution
    * Now utilises a class to help with readability
    * Fixed Office for more versions of Office
    * Removed the group check example as this increases the scripts run time by too much
    * encoding can be easily switch with a parameter 

#>

# [ Configuration ]

Write-Verbose "Version 6.0.0"
$folderLocation = Join-Path -Path $Env:appdata -ChildPath 'Microsoft\signatures'
$fileName = $SignatureName
$file  = Join-Path -Path $folderLocation -ChildPath $fileName

# [ Classes ]

class UserAccount {
    [string]$Name
    [string]$DistinguishedName
    [string]$UserPrincipalName
    [string]$DisplayName
    [string]$GivenName
    [string]$Initials
    [string]$Surname
    [string]$Description
    [string]$JobTitle
    [string]$Department
    [string]$Company
    [string]$EmailAddress
    [string]$StreetAddress
    [string]$City
    [string]$State
    [string]$PostalCode
    [string]$Country
    [string]$TelephoneNumber
    [string]$Mobile
    [string]$Pager
    [string]$Fax
    [string]$HomePhoneNumber
    [string]$OtherHomePhoneNumber
    [string]$HomeFax
    [string]$OtherFax
    [string]$IPPhone
    [string]$OtherIPPhone
    [string]$WebPage
    [string]$ExtensionAttribute1
    [string]$ExtensionAttribute2
    [string]$ExtensionAttribute3
    [string]$ExtensionAttribute4
    [string]$ExtensionAttribute5
    [string]$ExtensionAttribute6
    [string]$ExtensionAttribute7
    [string]$ExtensionAttribute8
    [string]$ExtensionAttribute9
    [string]$ExtensionAttribute10
    [string]$ExtensionAttribute11
    [string]$ExtensionAttribute12
    [string]$ExtensionAttribute13
    [string]$ExtensionAttribute14
    [string]$ExtensionAttribute15
}

# [ Functions ]

function Get-UserDetails {
    [CmdletBinding()]
    param ()

    try {
        $user = (([adsisearcher]"(&(objectCategory=User)(samaccountname=$env:username))").FindOne().Properties)
        Write-Verbose "User found for $($env:username)"

        $userAccount = New-Object UserAccount
        $userAccount.Name = $user['name'] -join ''
        $userAccount.DistinguishedName = $user['distinguishedname'] -join ''
        $userAccount.UserPrincipalName = $user['userprincipalname'] -join ''
        $userAccount.DisplayName = $user['displayname'] -join ''
        $userAccount.GivenName = $user['givenname'] -join ''
        $userAccount.Initials = $user['initials'] -join ''
        $userAccount.Surname = $user['sn'] -join ''
        $userAccount.Description = $user['description'] -join ''
        $userAccount.JobTitle = $user['title'] -join ''
        $userAccount.Department = $user['department'] -join ''
        $userAccount.EmailAddress = $user['mail'] -join ''
        $userAccount.StreetAddress = $user['streetaddress'] -join ''
        $userAccount.City = $user['l'] -join ''
        $userAccount.State = $user['st'] -join ''
        $userAccount.PostalCode = $user['postalcode'] -join ''
        $userAccount.Country = $user['c'] -join ''
        $userAccount.TelephoneNumber = $user['telephonenumber'] -join ''
        $userAccount.Mobile = $user['mobile'] -join ''
        $userAccount.Pager = $user['pager'] -join ''
        $userAccount.Fax = $user['facsimiletelephonenumber'] -join ''
        $userAccount.HomePhoneNumber = $user['homephone'] -join ''
        $userAccount.OtherHomePhoneNumber = $user['otherhomephone'] -join ''
        $userAccount.HomeFax = $user['homefax'] -join ''
        $userAccount.OtherFax = $user['otherfacsimiletelephonenumber'] -join ''
        $userAccount.IPPhone = $user['ipphone'] -join ''
        $userAccount.OtherIPPhone = $user['otheripphone'] -join ''

        $userAccount.Company = if ($null -eq $CompanyName) { $user['company'] } else { $CompanyName }
        $userAccount.WebPage = if ($null -eq $Website) { $user['wWWHomePage'] } else { $Website }
        
        $userAccount.ExtensionAttribute1 = $user['extensionattribute1'] -join ''
        $userAccount.ExtensionAttribute2 = $user['extensionattribute2'] -join ''
        $userAccount.ExtensionAttribute3 = $user['extensionattribute3'] -join ''

        return $userAccount
    } catch {
        Write-Error "Unable to query Active Directory for user information. Details: $($_.Exception.Message)"
        exit 1
    }
}

function Set-RegistryEntries {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Version,
        [Parameter(Mandatory = $true)]
        [string] $SignatureName
    )

    $regPathGeneral = "HKCU:\Software\Microsoft\Office\$Version\Common\General"
    $regPathMailSettings = "HKCU:\Software\Microsoft\Office\$Version\Common\MailSettings"
    $regPathOutlookSetup = "HKCU:\Software\Microsoft\Office\$Version\Outlook\Setup"

    Write-Verbose "Setting registry keys"
    New-ItemProperty -Path $regPathGeneral -Name "Signatures" -Value "signatures" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPathMailSettings -Name "NewSignature" -Value $SignatureName -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPathMailSettings -Name "ReplySignature" -Value $SignatureName -PropertyType String -Force | Out-Null

    Write-Verbose "Removing First-Run registry key"
    Remove-ItemProperty -Path $regPathOutlookSetup -Name "First-Run" -ErrorAction SilentlyContinue | Out-Null
}

function Get-OfficeVersion {
    [CmdletBinding()]
    param ()

    Write-Verbose "Getting office version"
    $officeVersions = @("16.0", "15.0", "14.0")
    $basePath = "HKCU:\Software\Microsoft\Office\"

    foreach ($version in $officeVersions) {
        $path = $basePath + $version + "\Outlook"
        if (Test-Path -Path $path) {
            Write-Verbose "Office Version: $version"
            return $version
        }
    }

    Write-Error "No compatible version of Microsoft Office found."
    exit 1
}

function Get-SignatureFromWordTemplate {
    [CmdletBinding()]
    param (
        [UserAccount]$user,
        [string]$TemplateFile,
        [string]$OutputPath,
        [string]$SignatureName
    )

    Write-Verbose "Generating signatures from Word template"

    try {
        # Check if Word template exists
        if (-not (Test-Path -Path $TemplateFile)) {
            Write-Error "Word template file not found at: $TemplateFile"
            exit 1
        }

        # Create Word Application object
        Write-Verbose "Starting Word application"
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $false
        $Word.DisplayAlerts = 0  # Disable alerts

        # Open the template document
        Write-Verbose "Opening template document: $TemplateFile"
        $Document = $Word.Documents.Open($TemplateFile)

        # Build address string
        $addressParts = @()
        if ($user.StreetAddress) { $addressParts += $user.StreetAddress }
        if ($user.City) { $addressParts += $user.City }
        if ($user.State) { $addressParts += $user.State }
        if ($user.PostalCode) { $addressParts += $user.PostalCode }
        $fullAddress = $addressParts -join ", "

        # Define replacement mappings
        $replacements = @{
            '{{DisplayName}}' = $user.DisplayName
            '{{JobTitle}}' = $user.JobTitle
            '{{Company}}' = $user.Company
            '{{Email}}' = $user.EmailAddress
            '{{Phone}}' = $user.TelephoneNumber
            '{{Mobile}}' = $user.Mobile
            '{{Website}}' = $user.WebPage
            '{{Address}}' = $fullAddress
            '{{Department}}' = $user.Department
            '{{Fax}}' = $user.Fax
            '{{GivenName}}' = $user.GivenName
            '{{Initials}}' = $user.Initials
            '{{Surname}}' = $user.Surname
            '{{Description}}' = $user.Description
        
        }

        # Replace placeholders in the document
        Write-Verbose "Replacing placeholders with user data"
        foreach ($placeholder in $replacements.Keys) {
            $value = $replacements[$placeholder]
            if ([string]::IsNullOrEmpty($value)) {
                $value = ""  # Replace with empty string if no value
            }
            
            $FindReplace = $Word.Selection.Find
            $FindReplace.Text = $placeholder
            $FindReplace.Replacement.Text = $value
            $FindReplace.Forward = $true
            $FindReplace.Wrap = 1  # wdFindContinue
            $FindReplace.Format = $false
            $FindReplace.MatchCase = $false
            $FindReplace.MatchWholeWord = $false
            $FindReplace.MatchWildcards = $false
            $FindReplace.MatchSoundsLike = $false
            $FindReplace.MatchAllWordForms = $false
            $FindReplace.Execute($placeholder, $false, $false, $false, $false, $false, $true, 1, $false, $value, 2)
        }

        # Save as RTF
        $rtfPath = Join-Path -Path $OutputPath -ChildPath "$SignatureName.rtf"
        Write-Verbose "Saving RTF signature: $rtfPath"
        $Document.SaveAs($rtfPath, 6)  # 6 = wdFormatRTF

        # Save as HTML
        $htmlPath = Join-Path -Path $OutputPath -ChildPath "$SignatureName.htm"
        Write-Verbose "Saving HTML signature: $htmlPath"
        $Document.SaveAs($htmlPath, 8)  # 8 = wdFormatHTML

        # Save as plain text
        $txtPath = Join-Path -Path $OutputPath -ChildPath "$SignatureName.txt"
        Write-Verbose "Saving TXT signature: $txtPath"
        $Document.SaveAs($txtPath, 2)  # 2 = wdFormatText

        # Close document and quit Word
        $Document.Close()
        $Word.Quit()

        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        Write-Verbose "Signature files generated successfully"
        return $true

    } catch {
        Write-Error "Error generating signatures from Word template: $($_.Exception.Message)"
        
        # Cleanup in case of error
        if ($Document) { 
            try { $Document.Close() } catch { }
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) } catch { }
        }
        if ($Word) { 
            try { $Word.Quit() } catch { }
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) } catch { }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        exit 1
    }
}

function Get-SignatureHTML {
    [CmdletBinding()]
    param (
        [UserAccount]$user
    )

    Write-Verbose "Building HTML from template"

    try {
        $templatePath = Join-Path -Path $PSScriptRoot -ChildPath "template.html"
        if (-not (Test-Path -Path $templatePath)) {
            Write-Error "HTML template file not found at: $templatePath"
            exit 1
        }
        
        $template = Get-Content -Path $templatePath -Raw
        
        # Use Invoke-Expression to evaluate the template with variables
        $expandedTemplate = $ExecutionContext.InvokeCommand.ExpandString($template)
        
        return $expandedTemplate
    } catch {
        Write-Error "Error reading or processing HTML template: $($_.Exception.Message)"
        exit 1
    }
}

function Get-SignaturePlainText {
    [CmdletBinding()]
    param (
        [UserAccount]$user
    )

    Write-Verbose "Building Plain Text from template"

    try {
        $templatePath = Join-Path -Path $PSScriptRoot -ChildPath "template.txt"
        if (-not (Test-Path -Path $templatePath)) {
            Write-Error "Plain text template file not found at: $templatePath"
            exit 1
        }
        
        $template = Get-Content -Path $templatePath -Raw
        
        # Use Invoke-Expression to evaluate the template with variables
        $expandedTemplate = $ExecutionContext.InvokeCommand.ExpandString($template)
        
        return $expandedTemplate
    } catch {
        Write-Error "Error reading or processing plain text template: $($_.Exception.Message)"
        exit 1
    }
}

# [ CSS ]

$style = 
@"
<style>
p, table, a, span { 
    font-family: Arial, Helvetica, sans-serif;
    font-size:  12pt;
}
</style>
"@

# [ Execution ]

$officeVersion = Get-OfficeVersion
$user = Get-UserDetails

# If the folder does not exist create it
if (-not (Test-Path -Path $folderLocation)) {
    try {
        Write-Verbose "Creating Directory: $folderLocation"
        New-Item -ItemType directory -Path $folderLocation
    } catch {
        Write-Error "Error: Unable to create the signatures folder. Details: $($_.Exception.Message)"
        exit
    }
}

if ($UseWordTemplate) {
    # Use Word template method
    Write-Verbose "Using Word template method"
    
    # Determine template filename
    if ([string]::IsNullOrWhiteSpace($UseWordTemplate)) {
        $templateFileName = "signature-template.docx"
    } else {
        $templateFileName = $UseWordTemplate
    }
    
    $wordTemplatePath = Join-Path -Path $PSScriptRoot -ChildPath $templateFileName
    
    if (-not (Test-Path -Path $wordTemplatePath)) {
        Write-Error "Word template file '$templateFileName' not found in script directory. Please create the template file or use the standard method without -UseWordTemplate parameter."
        exit 1
    }
    
    # Check if Word is available
    try {
        $testWord = New-Object -ComObject Word.Application -ErrorAction Stop
        $testWord.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($testWord) | Out-Null
        Write-Verbose "Microsoft Word is available"
    } catch {
        Write-Error "Microsoft Word is not available or not properly installed. Cannot use Word template method."
        exit 1
    }
    
    # Generate signatures using Word template
    $success = Get-SignatureFromWordTemplate -user $user -TemplateFile $wordTemplatePath -OutputPath $folderLocation -SignatureName $fileName
    
    if ($success) {
        Write-Verbose "Word template processing completed successfully"
    }
    
} else {
    # Use original method with HTML/TXT templates
    Write-Verbose "Using standard template method"
    
    $htmlSignature = Get-SignatureHTML -user $user
    $plainTextSignature = Get-SignaturePlainText -user $user

    # Save the HTML to the signature file
    try {
        Write-Verbose "Saving HTML signature"
        $style + $htmlSignature | Out-File -FilePath "$file.htm" -Encoding $Encoding
    } catch {
        Write-Error "Error: Unable to save the HTML signature file. Details: $($_.Exception.Message)"
        exit 1
    }

    # Output the text to the signatures folder
    try {
        Write-Verbose "Saving txt signature"
        $plainTextSignature | out-file "$file.txt" -encoding $Encoding
    } catch {
        Write-Error "Error: Unable to save the text signature file. Details: $($_.Exception.Message)"
        exit 1
    }
}

Set-RegistryEntries -Version $officeVersion -SignatureName $fileName