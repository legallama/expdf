# Email to PDF Exporter - Outlook Add-in

An Office Add-in for Outlook that exports emails to PDF with styling preserved. Export single emails or entire conversation chains with each message on a separate page.

## Features

- Export individual emails to PDF
- Export entire email conversations (each message on separate page)
- Preserve email styling and formatting
- Works on Outlook Web and Outlook Desktop
- Clean, professional PDF layout
- Includes all email metadata (From, To, CC, Date, Subject)

## Prerequisites

- Microsoft Outlook (Desktop or Web)
- Node.js 14+ (for development)
- Office 365 or Microsoft account

## Installation

### Option 1: Sideload (Development/Testing)

1. **Update manifest.xml**
   - Replace the URLs with your server location:
     - `functionfile` - typically `https://yourserver.com/functions/functions.html`
     - `taskpaneUrl` - typically `https://yourserver.com/taskpane.html`

2. **Host the files**
   - Upload `taskpane.html`, `taskpane.js`, and `taskpane.css` to your web server

3. **Sideload in Outlook Web Access**
   - Go to outlook.office.com
   - Click Settings → Apps → My add-ins → Get Add-ins → My add-ins
   - Select "Add a custom add-in" → "Add from File"
   - Choose your manifest.xml file

4. **Sideload in Outlook Desktop**
   - Windows: File → Get Add-ins → My add-ins → Upload My Add-in
   - Mac: Get Add-ins → Upload My Add-in

### Option 2: Office Store Submission

1. **Prepare for submission**
   - Update manifest.xml with your company details
   - Create app icons (16x16, 32x32, 80x80)
   - Update app name and description

2. **Test thoroughly**
   - Test on both Outlook Web and Desktop
   - Verify PDF export works on various email types
   - Test conversation exports

3. **Submit to Office Store**
   - Go to Partner Center
   - Create App submission
   - Upload manifest and supporting materials
   - Complete validation process

## Development

### Setup

```bash
npm install
```

### File Structure

- `manifest.xml` - Add-in configuration
- `taskpane.html` - UI interface
- `taskpane.js` - Core logic for email extraction and PDF generation
- `taskpane.css` - Styling
- `package.json` - Dependencies

### Key Technologies

- **Office JavaScript API** - Access Outlook email data
- **html2pdf.js** - Generate PDFs from HTML
- **Segoe UI Font** - Professional Microsoft styling

## Usage

1. Open an email in Outlook
2. Click the "Export to PDF" button in the ribbon (Outlook Desktop) or task pane icon (Outlook Web)
3. Select export option:
   - **Current Email Only** - Export just this email
   - **Entire Conversation** - Export all messages in thread (each on separate page)
4. Click "Export to PDF"
5. File downloads automatically

## Limitations

- Full conversation thread extraction requires server-side processing
- Attachments are not included in PDF (by design)
- External images may not load in PDF due to CORS restrictions
- Large emails may take a moment to process

## Troubleshooting

### Add-in not appearing
- Clear browser cache and refresh
- Ensure manifest.xml URLs are correct and accessible
- Check browser console for errors

### PDF not downloading
- Check popup blocker settings
- Verify html2pdf.js library is loading (check console)
- Try a different email to isolate the issue

### Styling not preserved
- Some styles may not render in PDF (limitations of html2pdf)
- Inline styles are more reliable than CSS classes

## Manifest Configuration

Key settings in manifest.xml:

- `Id` - Unique GUID for your add-in
- `Version` - Update when releasing new versions
- `ProviderName` - Your company/developer name
- `ExtensionPoint` - Controls where add-in appears in UI
- `Resources` - Icons, URLs, and localized strings

## Security Considerations

- Add-in runs in isolated sandbox environment
- Email content never sent to external servers (processes locally)
- Use HTTPS for all hosted files
- Validate manifest against Microsoft's schema

## Support & Updates

For issues or feature requests, contact the development team.

## License

MIT License
