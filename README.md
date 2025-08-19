# Email Voice Assistant

An advanced HTA application for automated email monitoring and voice responses.

## Features

- 🔍 Smart Email Monitoring with configurable intervals (10s-5min)
- ⏰ Flexible Time Windows (15min-24hrs) for email detection
- 📧 Multi-Account Support with Gmail account detection
- 🎯 Advanced Distribution List (DL) filtering
- 🔊 Automated voice responses (TTS or static audio)
- 🔄 Bulletproof deduplication system
- 📊 Enhanced logging and diagnostics

## Requirements

- Windows 10/11
- Outlook Desktop (classic/Win32) application
- Active email account configured in Outlook

## Quick Start

1. Double-click `OutlookVoiceAgent.hta` or use `start_agent.bat`
2. Configure settings in the UI:
   - Distribution List address
   - Keywords for monitoring
   - Reply mode and voice options
3. Click "Start Monitoring"
4. Use "Scan Now" for immediate testing

## Configuration Guide

### Basic Settings
- **DL Address**: Enter the Distribution List email or display name
- **Keywords**: Add comma-separated trigger words
- **Subject Filter**: Optional additional keyword filter
- **Poll Frequency**: Select check interval (10s to 5min)
- **Voice Mode**: Choose TTS or static audio file

### Advanced Options
- **Reply Mode**: Select Reply, ReplyAll, or NewToDL
- **Lookback Window**: Set initial scan period
- **Account Selection**: Choose specific email account

## Troubleshooting

### Common Issues

1. **COM Errors**
   - Ensure Outlook is running
   - Check Trust Center settings
   
2. **No Matches Found**
   - Verify DL address format
   - Check keyword spelling
   - Confirm account selection

3. **Account Issues**
   - Use "Select Inbox" button
   - Try "Refresh Folders"
   - Verify account access

### Debug Features

- Use "Copy Log" for detailed diagnostics
- Check processing summaries
- Review email matching details

## Security Notes

- Application runs with elevated privileges
- Only run trusted code
- Use Task Scheduler for secure autostart

## Task Scheduler Setup

1. Open Task Scheduler
2. Create new task:
   - Trigger: At log on
   - Action: mshta.exe
   - Arguments: "C:\Path\OutlookVoiceAgent.hta"
   - Optional: 30s delay for Outlook startup

## Uninstallation

1. Stop the application
2. Delete the application files
3. Remove Task Scheduler entry if configured

## Support

For issues or questions, please create a GitHub issue.