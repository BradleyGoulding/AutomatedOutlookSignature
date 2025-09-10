# Word Template Signature

This document serves as a template for generating Outlook signatures. The PowerShell script will:

1. Open this Word document
2. Replace placeholder text with actual user data
3. Save as RTF and HTML formats for Outlook

## Template Structure

The Word document should contain placeholders like:
- {{DisplayName}} - Will be replaced with user's display name
- {{JobTitle}} - Will be replaced with user's job title
- {{Company}} - Will be replaced with company name
- {{Email}} - Will be replaced with email address
- {{Phone}} - Will be replaced with phone number
- {{Mobile}} - Will be replaced with mobile number
- {{Website}} - Will be replaced with website URL
- {{Address}} - Will be replaced with full address

## Instructions

1. Create a Word document named "signature-template.docx"
2. Design your signature layout with formatting, images, tables, etc.
3. Use the placeholder text {{FieldName}} where you want dynamic content
4. The script will handle the rest automatically

This approach provides:
- Rich formatting capabilities
- Consistent branding
- Professional appearance
- Easy template maintenance

## Samples

Here are various ways to run the script with different options:

### Basic Usage

```powershell
# Use default Word template (template.docx)
.\Set-OutlookSignature.Ps1 -UseWordTemplate ""

# Use specific Word template
.\Set-OutlookSignature.Ps1 -UseWordTemplate "AEE Primary.docx"

# Use standard HTML/TXT templates (no Word)
.\Set-OutlookSignature.Ps1
```

### With Verbose Output

```powershell
# Default template with verbose logging
.\Set-OutlookSignature.Ps1 -UseWordTemplate "" -Verbose

# Specific template with verbose logging
.\Set-OutlookSignature.Ps1 -UseWordTemplate "AEE Primary.docx" -Verbose

# Standard templates with verbose logging
.\Set-OutlookSignature.Ps1 -Verbose
```

### Custom Signature Names

```powershell
# Corporate signature using default Word template
.\Set-OutlookSignature.Ps1 -UseWordTemplate "" -SignatureName "Corporate"

# Executive signature using specific template
.\Set-OutlookSignature.Ps1 -UseWordTemplate "Executive Template.docx" -SignatureName "Executive"

# Department-specific signatures
.\Set-OutlookSignature.Ps1 -UseWordTemplate "Marketing.docx" -SignatureName "Marketing Team"
.\Set-OutlookSignature.Ps1 -UseWordTemplate "Sales.docx" -SignatureName "Sales Team"
```

### Override Company Information

```powershell
# Override company name and website
.\Set-OutlookSignature.Ps1 -UseWordTemplate "AEE Primary.docx" -CompanyName "Acme Corporation" -Website "www.acme.com"

# Override with verbose output
.\Set-OutlookSignature.Ps1 -UseWordTemplate "" -CompanyName "Custom Corp" -Website "custom.com" -Verbose

# Complete customization
.\Set-OutlookSignature.Ps1 -UseWordTemplate "Custom.docx" -SignatureName "Custom Sig" -CompanyName "My Company" -Website "mysite.com" -Verbose
```

### Different Encoding Options

```powershell
# UTF-8 encoding
.\Set-OutlookSignature.Ps1 -UseWordTemplate "AEE Primary.docx" -Encoding "utf8"

# Unicode encoding (default)
.\Set-OutlookSignature.Ps1 -UseWordTemplate "AEE Primary.docx" -Encoding "unicode"

# ASCII encoding
.\Set-OutlookSignature.Ps1 -UseWordTemplate "AEE Primary.docx" -Encoding "ascii"
```

### Advanced Scenarios

```powershell
# Multiple signatures for different purposes
.\Set-OutlookSignature.Ps1 -UseWordTemplate "Internal.docx" -SignatureName "Internal Communications"
.\Set-OutlookSignature.Ps1 -UseWordTemplate "External.docx" -SignatureName "External Communications"
.\Set-OutlookSignature.Ps1 -UseWordTemplate "Marketing.docx" -SignatureName "Marketing Campaigns"

# Test run with all options
.\Set-OutlookSignature.Ps1 -UseWordTemplate "Test Template.docx" -SignatureName "Test Signature" -CompanyName "Test Company" -Website "test.com" -Encoding "utf8" -Verbose

# Fallback to standard templates if Word template fails
.\Set-OutlookSignature.Ps1 -SignatureName "Fallback" -Verbose
```

### Batch Processing Examples

```powershell
# Create multiple signatures from different templates
$templates = @(
    @{Template="Executive.docx"; Name="Executive"; Company="Acme Corp"},
    @{Template="Manager.docx"; Name="Manager"; Company="Acme Corp"},
    @{Template="Employee.docx"; Name="Standard"; Company="Acme Corp"}
)

foreach ($config in $templates) {
    .\Set-OutlookSignature.Ps1 -UseWordTemplate $config.Template -SignatureName $config.Name -CompanyName $config.Company -Verbose
}
```

### Troubleshooting Commands

```powershell
# Test if Word is available
try {
    $word = New-Object -ComObject Word.Application
    Write-Host "✅ Word is available" -ForegroundColor Green
    $word.Quit()
} catch {
    Write-Host "❌ Word not available: $($_.Exception.Message)" -ForegroundColor Red
}

# Check if template file exists
$templatePath = ".\AEE Primary.docx"
if (Test-Path $templatePath) {
    Write-Host "✅ Template found: $templatePath" -ForegroundColor Green
} else {
    Write-Host "❌ Template not found: $templatePath" -ForegroundColor Red
}

# List available template files
Get-ChildItem -Path . -Filter "*.docx" | Select-Object Name, LastWriteTime
```
