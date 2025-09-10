# Testing Word Template Functionality

## Quick Test Script

Here's a simple PowerShell snippet to test if Word automation is working:

```powershell
# Test if Word is available
try {
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    Write-Host "✅ Microsoft Word is available" -ForegroundColor Green
    $Word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
} catch {
    Write-Host "❌ Microsoft Word is not available: $($_.Exception.Message)" -ForegroundColor Red
}
```

## Create Sample Word Template

1. Open Microsoft Word
2. Create a new document
3. Add this content:

```
{{DisplayName}}
{{JobTitle}}

{{Company}}
{{Address}}

Email: {{Email}}
Phone: {{Phone}}
Mobile: {{Mobile}}
Website: {{Website}}
```

4. Format it nicely (fonts, colors, layout)
5. Save as `signature-template.docx` in your script directory

## Run the Script

```powershell
# Test with Word template
.\Set-OutlookSignature.Ps1 -UseWordTemplate -Verbose

# Test with standard templates
.\Set-OutlookSignature.Ps1 -Verbose
```

## Expected Output

With `-UseWordTemplate`, you should see:
- `Signature.rtf` - Rich Text Format
- `Signature.htm` - HTML Format  
- `Signature.txt` - Plain Text Format

All three files will be generated from your Word template with proper formatting preserved.

## Troubleshooting

### "Word is not available"
- Ensure Microsoft Word is installed
- Try running PowerShell as Administrator
- Check if Word processes are running in Task Manager

### "Template file not found"
- Ensure `signature-template.docx` is in the same folder as the PowerShell script
- Check the filename is exactly `signature-template.docx`

### "Access denied" or COM errors
- Close any open Word documents
- Run PowerShell as Administrator
- Restart the computer if Word processes are stuck
