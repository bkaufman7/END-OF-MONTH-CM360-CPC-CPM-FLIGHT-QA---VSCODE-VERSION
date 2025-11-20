# CM360 End of Month Audit - Deployment Guide

## Prerequisites

1. **Google Account**: With access to Google Apps Script and Gmail
2. **Google Spreadsheet**: Target spreadsheet for the audit system
3. **Gmail Labels**: "CM360 QA" label for report identification
4. **Node.js**: For clasp CLI tool (optional but recommended)

## Setup Steps

### 1. Prepare Your Google Spreadsheet

Create a new Google Spreadsheet with the following sheets:

#### Required Sheets
- **Raw Data** - Headers: Network ID, Advertiser, Placement ID, Placement, Campaign, Placement Start Date, Placement End Date, Campaign Start Date, Campaign End Date, Ad, Impressions, Clicks, Report Date
- **Violations** - Headers: Network ID, Report Date, Advertiser, Campaign, Campaign Start Date, Campaign End Date, Ad, Placement ID, Placement, Placement Start Date, Placement End Date, Impressions, Clicks, CTR (%), Days Until Placement End, Flight Completion %, Days Left in the Month, CPC Risk, $CPC, $CPM, Issue Type, Details, Last Imp Change, Last Click Change, Owner (Ops)
- **Networks** - Headers: Network ID, Network Name, [additional owner/ops columns]
- **EMAIL LIST** - Column A: Email addresses for report recipients
- **Advertisers to ignore** - Column A: Advertiser names to exclude

### 2. Install Google Apps Script CLI (Optional)

```bash
npm install -g @google/clasp
clasp login
```

### 3. Create Apps Script Project

#### Option A: Manual Setup
1. Go to [script.google.com](https://script.google.com)
2. Create new project
3. Replace Code.gs content with your CM360 audit code
4. Update appsscript.json with required services

#### Option B: Using Clasp
```bash
# In your project directory
clasp create --type sheets --title "CM360 End of Month Audit"
clasp push
```

### 4. Enable Required APIs

In Google Apps Script Editor:
1. Click "Services" in left sidebar
2. Add:
   - **Gmail API** (v1)
   - **Drive API** (v3)

### 5. Configure Spreadsheet Binding

1. In Apps Script editor, go to Settings (gear icon)
2. Under "General settings", link to your target spreadsheet
3. Or use: `Resources > Cloud Platform project` to link existing sheet

### 6. Set Up Gmail Label

1. In Gmail, create label: "CM360 QA"
2. Apply this label to emails containing CM360 reports
3. Ensure CSV/ZIP attachments are present

### 7. Configure Network Mapping

In the **Networks** sheet:
- Column A: Network ID (numeric)
- Column B: Network Name  
- Columns P-S: Owner/Ops information (script looks for "ops" in header)

### 8. Add Email Recipients  

In **EMAIL LIST** sheet:
- Column A: Add email addresses for automated reports
- One email per row, starting from A2

### 9. Initial Authorization

1. Run from Apps Script editor: `authorizeMail_()` function
2. Grant all requested permissions
3. Verify test email is received

### 10. Set Up Automation

From the Google Sheets menu (after deployment):
1. **CM360 QA Tools** > **Create Daily Email Trigger (9am)**
2. This creates automatic daily summary emails

## Testing

### Manual Testing
1. **CM360 QA Tools** > **Pull Data** - Test Gmail import
2. **CM360 QA Tools** > **Run QA Only** - Test violation detection
3. **CM360 QA Tools** > **Send Email Only** - Test email reports

### Full Workflow Test
1. **CM360 QA Tools** > **Run It All** - Complete end-to-end test

## Configuration Options

### Time Zone
Update `appsscript.json`:
```json
{
  "timeZone": "America/New_York"
}
```

### Processing Limits (in Code.gs)
```javascript
const QA_CHUNK_ROWS = 3500;           // Rows per chunk
const QA_TIME_BUDGET_MS = 4.2 * 60 * 1000;  // Time per chunk
```

### Business Rules
```javascript
// Performance thresholds
PERFORMANCE_ALERT_CTR_THRESHOLD: 90    // CTR percentage
PERFORMANCE_ALERT_CPM_THRESHOLD: 10    // CPM dollar amount

// Cost thresholds  
HIGH_COST_CPC_THRESHOLD: 10           // CPC dollar amount
HIGH_COST_CPM_THRESHOLD: 10           // CPM dollar amount
```

## Monitoring and Maintenance

### Logs
- View execution logs in Apps Script editor
- Monitor email delivery success/failures
- Check chunked processing progress

### Performance
- Large datasets automatically use chunked processing
- State management enables resume after timeouts
- Lock management prevents concurrent execution conflicts

### Troubleshooting

#### Common Issues
1. **Permission Denied**: Re-run `authorizeMail_()` function
2. **No Data**: Check Gmail label "CM360 QA" exists and has attachments
3. **Email Failures**: Verify EMAIL LIST sheet has valid addresses
4. **Chunked Processing Stuck**: Clear QA state in DocumentProperties

#### Reset Options
- **Clear Violations**: Clears violation tracking data
- **Manual State Reset**: Delete DocumentProperties 'qa_progress_v2'
- **Cache Cleanup**: Hidden sheets are auto-managed

## Security Notes

- Script requires Gmail and Drive permissions
- Email addresses in EMAIL LIST receive sensitive audit data
- Sidecar spreadsheets contain historical violation tracking
- Regular cleanup of old cache data is automatic

## Support

For issues or modifications, reference the detailed code comments and Logger statements throughout the system.