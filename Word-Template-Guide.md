# Creating Your Word Template (signature-template.docx)

## Instructions

1. **Create a new Word document** and save it as `signature-template.docx` in the same folder as the PowerShell script.

2. **Design your signature** using Word's formatting tools:
   - Add your company logo
   - Set fonts, colors, and styling
   - Create tables for contact information
   - Add borders, lines, or other design elements

3. **Use placeholders** for dynamic content. The script will replace these with actual user data:

   ```
   {{DisplayName}} - User's full display name
   {{JobTitle}} - Job title/position
   {{Company}} - Company name
   {{Department}} - Department name
   {{Email}} - Email address
   {{Phone}} - Office phone number
   {{Mobile}} - Mobile phone number
   {{Fax}} - Fax number
   {{Website}} - Company website
   {{Address}} - Full address (street, city, state, postal code)
   ```

## Sample Template Layout

Here's an example of what your Word document might contain:

---

**{{DisplayName}}**
*{{JobTitle}}*

[INSERT COMPANY LOGO HERE]

**{{Company}}**
{{Address}}

📧 {{Email}}
📞 {{Phone}}
📱 {{Mobile}}
🌐 {{Website}}

---

## Advanced Features

- **Conditional Content**: If a user field is empty, the placeholder will be replaced with an empty string
- **Rich Formatting**: Use Word's full formatting capabilities (bold, italic, colors, tables, etc.)
- **Images**: Insert company logos, social media icons, or other graphics
- **Tables**: Create structured layouts for contact information
- **Hyperlinks**: The script will preserve hyperlink formatting for emails and websites

## File Formats Generated

The script will automatically generate three signature files from your Word template:
- **signature.rtf** - Rich Text Format (recommended for Outlook)
- **signature.htm** - HTML format (for web-based email clients)
- **signature.txt** - Plain text format (fallback)

## Usage

Run the script with the Word template option:
```powershell
.\Set-OutlookSignature.Ps1 -UseWordTemplate -Verbose
```

## Benefits of Word Templates

1. **WYSIWYG Design** - See exactly how your signature will look
2. **Rich Formatting** - Full Word formatting capabilities
3. **Company Branding** - Easy to incorporate logos and brand colors
4. **Professional Appearance** - Better than hand-coded HTML
5. **Easy Maintenance** - Non-technical users can update templates
6. **Consistent Output** - Same formatting across all three file types
