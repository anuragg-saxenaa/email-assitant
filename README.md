# Outlook Voice Agent

An HTA application that monitors your Outlook inbox for emails sent to a specific Distribution List, checks for configured keywords, and automatically responds with a voice note attachment.

## What It Does

- Polls your Outlook inbox in the background at configurable intervals
- Filters messages addressed to a specific Distribution List (DL)
- Scans subject and body for configured keywords
- When a match is found, automatically replies with:
  - A voice note attachment (either TTS-generated or a static audio file)
  - A short HTML message in the email body
- Marks processed emails with a category tag to avoid duplicate replies
- Skips typical auto-replies (Out of Office, undeliverable) and noreply senders

## Limitations

- Requires Outlook desktop application (Win32) to be open and running
- HTA runs with high privileges; only run code you trust
- TTS produces WAV files; if MP3 is required, use a static MP3 file instead
- Cannot process shared/team mailboxes unless the Inbox is accessible in your current Outlook profile
- Requires Windows 10/11 with Microsoft Outlook installed

## Quick Start

1. **Edit Configuration**:
   - Open `OutlookVoiceAgent.hta` in a text editor
   - Locate the configuration section at the top of the file
   - Set your DL address, keywords, and other parameters

2. **Run the Application**:
   - Double-click `OutlookVoiceAgent.hta` to launch
   - Click "Scan Now" to test immediately
   - Click "Minimize" to keep it running in the background

3. **Test**:
   - Send a test email to your configured DL with one of your keywords
   - The agent should detect it and send a voice reply

## Configuration Options

Edit these settings at the top of the HTA file:

- `DL_ADDRESS`: The Distribution List SMTP address or display name to match
- `KEYWORDS`: Array of case-insensitive keywords to trigger replies
- `POLL_MS`: Polling interval in milliseconds (default: 30000 = 30 seconds)
- `LOOKBACK_MINUTES`: How far back to check for emails when starting (default: 10)
- `USE_TTS`: Set to true to generate speech, false to use a static audio file
- `VOICE_TEXT`: The text to convert to speech when USE_TTS is true
- `STATIC_AUDIO_PATH`: Path to your MP3/WAV file when USE_TTS is false
- `REPLY_MODE`: "Reply" (to sender), "ReplyAll" (to all), or "NewToDL" (new email to DL)
- `PROCESS_CATEGORY`: Category label to mark processed emails
- `SKIP_SUBJECT_PATTERNS` and `SKIP_SENDER_PATTERNS`: Patterns to avoid replying to

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

### COM Errors or Outlook Connection Issues

- Ensure Outlook desktop application is running before starting the agent
- Check your Outlook Trust Center settings (File > Options > Trust Center > Trust Center Settings > Programmatic Access) and ensure it's not set to deny programmatic access
- Try running the HTA as administrator

### No Matches Found

- Verify your DL address is correct (SMTP address works best)
- Increase the LOOKBACK_MINUTES value
- Check that emails are actually in your Inbox (not in another folder)
- Verify your keywords are present in the subject or body

### Duplicate Replies

- Confirm the category is being applied correctly
- Ensure no other rules or scripts are also replying to the same emails
- Check if multiple instances of the agent are running

### TTS Not Working

- Ensure Microsoft SAPI is installed and working
- Try using a static audio file instead (set USE_TTS to false)

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
# email-assitant
