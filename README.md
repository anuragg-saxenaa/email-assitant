# Outlook Voice Agent

An advanced HTA application that monitors your Outlook inbox for emails sent to a specific Distribution List, checks for configured keywords, and automatically responds with a voice note attachment. Features a modern UI with comprehensive configuration options and robust duplicate prevention.

## What It Does

- **Smart Email Monitoring**: Polls your Outlook inbox with user-configurable intervals (10 seconds to 5 minutes)
- **Flexible Time Windows**: Configurable lookback window (15 minutes to 24 hours) for catching emails
- **Multi-Account Support**: Automatically detects Gmail accounts and supports manual inbox selection
- **Advanced Filtering**: Filters messages addressed to specific Distribution Lists with keyword matching
- **Intelligent Processing**: Scans subject and body for configured keywords with detailed logging
- **Automated Responses**: When matches are found, automatically replies with:
  - Voice note attachments (TTS-generated or static audio files)
  - Professional HTML message bodies
  - Configurable reply modes (Reply, Reply All, or New Email to DL)
- **Bulletproof Deduplication**: Dual-layer system prevents duplicate replies using both categories and ID tracking
- **Smart Filtering**: Automatically skips auto-replies, out-of-office messages, and noreply senders
- **Comprehensive Logging**: Detailed processing summaries with actionable suggestions for troubleshooting

## Limitations

- Requires Outlook desktop application (Win32) to be open and running
- HTA runs with high privileges; only run code you trust
- TTS produces WAV files; if MP3 is required, use a static MP3 file instead
- Cannot process shared/team mailboxes unless accessible in your current Outlook profile
- Requires Windows 10/11 with Microsoft Outlook installed
- Uses JScript engine for maximum compatibility (no modern JavaScript features)

## Quick Start

1. **Launch the Application**:
   - Double-click `OutlookVoiceAgent.hta` to launch
   - Or use the included `start_agent.bat` file
   - The application will automatically detect your email accounts

2. **Configure Settings** (No code editing required!):
   - **DL Email/Name**: Enter your Distribution List address or name
   - **Keywords**: Add comma-separated keywords to trigger responses
   - **Lookback Window**: Choose how far back to check emails (15 min to 24 hours)
   - **Poll Frequency**: Set how often to check for new emails (10 sec to 5 min)
   - **Reply Mode**: Choose Reply, Reply All, or New Email to DL
   - **Voice Settings**: Configure TTS text or static audio file path

3. **Start Monitoring**:
   - Click "Start Monitoring" to begin
   - Use "Scan Now" for immediate testing
   - Monitor the log for detailed processing information

4. **Test & Troubleshoot**:
   - Send a test email to your configured DL with keywords
   - Check the comprehensive log summary for any issues
   - Use "Copy Log" to export debug information if needed

## Configuration Options

All settings are now configurable through the modern UI interface:

### **User-Configurable Settings**
- **DL Email/Name**: Distribution List SMTP address or display name to match
- **Keywords**: Comma-separated, case-insensitive keywords to trigger replies
- **Subject Filter**: Optional additional subject keyword filter
- **Lookback Window**: How far back to check emails when starting
  - Options: 15 min, 30 min, 1 hour (default), 2 hours, 4 hours, 8 hours, 24 hours
- **Poll Frequency**: How often to check for new emails
  - Options: 10 sec, 15 sec, 30 sec (default), 1 min, 2 min, 5 min
- **Reply Mode**: How to respond to matched emails
  - "Reply" (to sender only), "Reply All" (to all recipients), "New Email to DL"
- **Voice Mode**: Choose between TTS generation or static audio file
- **TTS Text**: Custom text for speech synthesis
- **Audio File Path**: Path to static MP3/WAV file

### **Advanced Features**
- **Multi-Account Support**: Automatic Gmail detection with manual selection
- **Dual-Layer Deduplication**: Category-based + ID tracking prevents duplicates
- **Smart Skip Patterns**: Automatically avoids auto-replies and noreply senders
- **Comprehensive Logging**: Detailed processing summaries with troubleshooting suggestions

## Auto-Start with Windows (Task Scheduler)

To have the agent start automatically when you log in:

1. Open Task Scheduler (search for it in the Start menu)
2. Click "Create Basic Task"
3. Name it "Outlook Voice Agent" and click Next
4. Select "When I log on" and click Next
5. Select "Start a program" and click Next
6. For Program/script, enter: `mshta.exe`
7. For Add arguments, enter the full path to your HTA in quotes: `"C:\Path\To\OutlookVoiceAgent.hta"`
8. Click Next, then Finish
9. Optional: Right-click the new task, select Properties, and check "Run with highest privileges"
10. Optional: In the Triggers tab, edit the trigger to add a 30-second delay to let Outlook start first

Alternatively, you can use the included `start_agent.bat` file and add it to your Startup folder.

## Troubleshooting

The application now provides comprehensive diagnostic information to help resolve issues quickly.

### **Email Processing Issues**

**📧 No Emails Being Processed**:
- Check the **Email Processing Summary** at the end of each scan
- Look for specific reasons why emails were skipped
- Follow the **💡 Suggestions to Fix** provided in the log

**🎯 DL Matching Issues**:
- Verify your DL setting matches the email recipients exactly
- Try both SMTP address (user@domain.com) and display name ("Team Name")
- Check the detailed breakdown showing "Expected DL" vs "Actual recipients"

**🔍 Keyword Matching Issues**:
- Review the keyword list shown in processing logs
- Check if keywords appear in subject or body as displayed
- Consider adding more generic keywords or variations

### **Technical Issues**

**COM Errors or Outlook Connection**:
- Ensure Outlook desktop application is running before starting
- Check Outlook Trust Center settings (File > Options > Trust Center > Programmatic Access)
- Try running the HTA as administrator
- Use "Select Inbox" to choose the correct email account

**Multi-Account Issues**:
- Use "Refresh Folders" to explore available accounts
- Click "Select Inbox" to manually choose from detected accounts
- Check logs for Gmail account detection messages

**Time Window Issues**:
- Increase the Lookback Window if emails are older
- Check the debug logs showing time comparisons
- Verify emails are within the configured time range

**TTS/Audio Issues**:
- Ensure Microsoft SAPI is installed and working
- Try using a static audio file instead of TTS
- Check file paths are correct and accessible

### **Using the Enhanced Logging**

The application now provides detailed summaries after each scan:
- **📊 Results Overview**: Shows processed vs skipped counts
- **📋 Detailed Breakdown**: Lists each email with specific skip reasons
- **💡 Actionable Suggestions**: Provides exact steps to fix configuration issues
- **Copy Log**: Export complete diagnostic information for support

## Security Notes

- HTAs run with the same permissions as the current user
- The application does not make any external network calls
- No data is sent outside of your local Outlook environment

## Uninstalling

To uninstall the agent:

1. Close the HTA if it's running
2. Delete the HTA file and associated batch file
3. Remove any Task Scheduler tasks you created for auto-start

## License

This software is provided under the MIT License. See the LICENSE.txt file for details.
## Recent Updates

### Version 2.0 Features
- **🎛️ Dynamic UI Configuration**: No more code editing required
- **📧 Multi-Account Support**: Automatic Gmail detection and manual selection
- **🔒 Enhanced Deduplication**: Bulletproof duplicate prevention system
- **📊 Comprehensive Logging**: Detailed processing summaries with troubleshooting guidance
- **⏰ Flexible Timing**: User-configurable lookback windows and poll frequencies
- **🛠️ Robust Error Handling**: Graceful recovery with detailed error reporting
- **📋 Improved Log Management**: Enhanced copy functionality compatible with HTA environment
- **🎯 Smart Diagnostics**: Actionable suggestions for configuration issues
