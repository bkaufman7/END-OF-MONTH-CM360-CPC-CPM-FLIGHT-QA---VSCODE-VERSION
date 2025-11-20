<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->
- [x] Verify that the copilot-instructions.md file in the .github directory is created.

- [x] Clarify Project Requirements
	<!-- Google Apps Script project for CM360 End of Month Audit with Gmail integration and automated reporting -->

- [x] Scaffold the Project
	<!--
	Created Google Apps Script project structure with:
	- Code.gs (main script file)
	- appsscript.json (configuration)
	- package.json (Node.js dependencies)
	- README.md (documentation)
	- DEPLOYMENT.md (setup guide)
	- .clasp.json.template (CLI configuration)
	- .gitignore (version control)
	-->

- [x] Customize the Project
	<!--
	Implemented CM360 End of Month Audit system with:
	- Gmail integration for CSV/ZIP processing
	- Chunked execution with state management  
	- Performance alerts and violation tracking
	- Low-priority classification system
	- Owner mapping and email reporting
	- Complete user's existing functionality preserved
	-->

- [x] Install Required Extensions
	<!-- No VS Code extensions required for Google Apps Script project -->

- [x] Compile the Project
	<!--
	Google Apps Script project configured with:
	- Gmail API v1 enabled
	- Drive API v3 enabled  
	- V8 runtime
	- Stackdriver logging
	- Eastern timezone configuration
	-->

- [x] Create and Run Task
	<!-- Google Apps Script projects run in cloud, no local tasks needed -->

- [x] Launch the Project
	<!--
	Project launches via Google Apps Script editor:
	1. Upload to script.google.com
	2. Bind to target Google Spreadsheet
	3. Run authorization functions
	4. Set up daily triggers
	-->

- [x] Ensure Documentation is Complete
	<!--
	Documentation complete:
	- README.md: Comprehensive feature overview and setup
	- DEPLOYMENT.md: Step-by-step deployment guide
	- copilot-instructions.md: Project configuration notes
	- complete-implementation.js: Full code reference
	-->

# CM360 End of Month Audit System

## Project Overview
Google Apps Script-based automation for CM360 end-of-month auditing with Gmail integration, violation tracking, and email reporting.

## Key Features
- Automated Gmail CSV/ZIP processing
- Chunked execution for large datasets (3500 rows/chunk)
- Performance spike alerts (pre-15th)
- Monthly violation summaries (post-15th)  
- Low-priority placement classification
- Historical violation tracking
- Owner-based reporting

## Next Steps
1. Upload Code.gs to Google Apps Script
2. Configure target Google Spreadsheet binding
3. Enable Gmail and Drive APIs
4. Set up required sheets and data
5. Run initial authorization
6. Configure daily automation triggers

## Files Ready for Deployment
- ✅ Code.gs - Core script implementation  
- ✅ appsscript.json - API configuration
- ✅ complete-implementation.js - Full code reference
- ✅ README.md - System documentation
- ✅ DEPLOYMENT.md - Setup instructions