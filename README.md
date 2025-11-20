# CM360 End of Month Audit System

A Google Apps Script-based system for automated CM360 end-of-month auditing with Gmail integration, violation tracking, and email reporting.

## Features

### 🔄 Automated Data Processing
- **Gmail Integration**: Automatically processes CM360 CSV/ZIP reports from Gmail attachments
- **Chunked Execution**: Handles large datasets with intelligent chunking (3500 rows per chunk)
- **State Management**: Resumes processing where it left off using DocumentProperties

### 📊 QA Engine
- **Billing Risk Detection**: Identifies CPC billing risks, expired placements, and cost anomalies
- **Delivery Issues**: Detects post-flight activity and delivery problems  
- **Performance Alerts**: Monitors CTR and CPM thresholds with spike detection
- **Low-Priority Classification**: Automatically identifies tracking pixels and non-creative placements

### 👥 Owner Management
- **Network-Based Mapping**: Resolves placement owners using Networks sheet data
- **Advertiser Normalization**: Handles advertiser name variations and matching
- **Ignore Lists**: Filters out specified advertisers from processing

### 📧 Email Reporting
- **Pre-15th Alerts**: Performance spike notifications before mid-month
- **Monthly Summaries**: Comprehensive violation reports after 15th
- **Size-Controlled**: Intelligent email sizing to prevent delivery issues
- **XLSX Exports**: Attached spreadsheet reports for detailed analysis

### 🏷️ Violation Tracking
- **Historical Changes**: Tracks impression/click changes over time
- **Sidecar Storage**: Uses separate spreadsheet for violation cache
- **Change Detection**: Identifies new/modified violations for alerts
- **Cleanup**: Automatic pruning of old tracking data

## Google Sheets Structure

### Required Sheets
- **Raw Data**: Imported CM360 report data
- **Violations**: QA results and issue tracking
- **Networks**: Network ID to owner mapping
- **EMAIL LIST**: Recipients for automated reports
- **Advertisers to ignore**: Excluded advertiser list

### Auto-Created Sheets  
- **_Perf Alert Cache**: Performance alert change tracking
- **_Violation Change Cache**: Historical violation data (sidecar)

## Menu Functions

- **Run It All**: Complete workflow - import, QA, alerts, summary
- **Pull Data**: Import CM360 reports from Gmail only
- **Run QA Only**: Process existing data for violations
- **Send Email Only**: Generate and send summary reports
- **Clear Violations**: Reset violations sheet
- **Authorize Email**: One-time MailApp authorization
- **Create Daily Email Trigger**: Set up 9am daily trigger

## Configuration

### Gmail Labels
- Uses label "CM360 QA" for report identification
- Processes CSV and ZIP attachments from current day

### Processing Limits
- **Chunk Size**: 3500 rows per execution
- **Time Budget**: 4.2 minutes per chunk  
- **Auto-Resume**: Schedules continuation triggers automatically

### Business Rules
- **Performance Threshold**: CTR ≥ 90% & CPM ≥ $10
- **Cost Thresholds**: CPC/CPM > $10 for alerts
- **Billing Risk**: Clicks > impressions detection
- **Stale Metrics**: Configurable days threshold (default 7)

## Setup Instructions

1. **Create Google Spreadsheet**: Bound to this Apps Script project
2. **Set up Sheets**: Create required sheets with proper headers
3. **Configure Networks**: Map Network IDs to owner information  
4. **Email Recipients**: Add email addresses to EMAIL LIST sheet
5. **Gmail Labels**: Ensure "CM360 QA" label exists in Gmail
6. **Authorization**: Run "Authorize Email" from menu
7. **Daily Triggers**: Use "Create Daily Email Trigger" for automation

## Technical Details

### State Management
- Uses DocumentProperties for QA session tracking
- Implements chunked processing with resume capability
- Manages trigger scheduling for large dataset processing

### Error Handling  
- Retry logic with exponential backoff
- Lock management for concurrent execution prevention
- Graceful degradation for quota limits

### Performance Optimization
- Batch operations for large datasets
- Memory-efficient processing patterns
- Intelligent caching for violation tracking

## Monitoring

The system provides detailed logging for:
- Processing progress and timing
- Email delivery status
- Error conditions and recovery
- Data validation results

## Made by Platform Solutions Automation (BK)