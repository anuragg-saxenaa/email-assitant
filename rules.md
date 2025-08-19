## Goal
Create an HTA that:
- Polls Outlook Inbox in the background.
- Filters messages addressed to a specific *Distribution List (DL)*.
- Scans *subject/body* for configured *keywords*.
- If matched, *auto-replies* (Reply / ReplyAll / or New mail to the DL) with:
  - A *predefined voice note* attached: either
    - generated on the fly via *Windows SAPI TTS* (WAV), or
    - a *static* MP3/WAV from disk.
  - A short canned HTML message in the email body.
- Dedupes with a *Category* tag (e.g., AI-VoiceReplied) and dual-layer ID tracking.
- Skips typical auto-replies (OOO, undeliverable) and noreply senders.
- Provides a user-friendly interface for dynamic configuration without code editing.

## Output (files)
- OutlookVoiceAgent.hta — self-contained UI + logic (HTML + inline JS/CSS).
- README.md — setup, config, and Task Scheduler instructions.
- start_agent.bat — optional helper to start via mshta.exe.
- license.txt — MIT License.

## HTA requirements
- Pure HTML + JS (JScript) using *ActiveX/COM*:
  - Outlook.Application (MAPI) to read inbox + send replies.
  - SAPI.SpVoice + SAPI.SpFileStream for optional TTS WAV generation.
  - Scripting.FileSystemObject for temp file handling.
- Runs on Windows 10/11 with *Outlook desktop (classic/Win32)* open.
- Poll interval configurable (default: 30s).
- Startup lookback window configurable (default: 60 min).
- Robust *error handling*: never crashes; log pane shows errors/status.
- *No external network calls* required for core function.
- Code must be *clean, commented, and production-ready*.
- Compatible with older JScript engines (no ES5+ methods like .trim(), .map()).

## UI Features (Dynamic Configuration)
- **Configuration Panel**: User-friendly input fields for all settings:
  - DL Email/Name input field
  - Keywords textarea (comma-separated)
  - Subject filter (optional)
  - Reply mode dropdown (Reply/ReplyAll/NewToDL)
  - Voice mode selector (TTS/Static Audio)
  - TTS text input / Static audio file path
- **Status Panel**: Real-time monitoring information
- **Log Area**: Scrollable log with timestamps and color-coded messages
- **Control Buttons**:
  - Start/Stop Monitoring toggle
  - Scan Now (manual scan)
  - Refresh Folders (explore available inboxes)
  - Copy Log (copy all log content to clipboard)
  - Clear Log (clear log area)
  - Select Inbox (choose from multiple email accounts)
  - Minimize/Exit

## Enhanced Logic Features
- **Multi-Account Support**:
  - Automatic Gmail account detection
  - Manual inbox selection from available accounts
  - Detailed account exploration and debugging
- **Dual-Layer Deduplication**:
  - Primary: Outlook category-based tracking
  - Secondary: Email ID tracking using EntryID or sender+subject+time
  - Prevents duplicate replies even if category system fails
- **Comprehensive Debugging**:
  - Detailed step-by-step processing logs
  - Email filtering decision tracking
  - TTS/attachment creation debugging
  - Reply sending process monitoring
- **Robust Error Handling**:
  - Graceful fallbacks for all operations
  - Detailed error logging with context
  - Continue operation even if individual steps fail

## Config (Dynamic UI Inputs)
All parameters now configurable through UI without code editing:
- **DL_ADDRESS** (input field) — the DL SMTP or display name to match.
- **KEYWORDS** (textarea) — case-insensitive subject/body triggers.
- **SUBJECT_FILTER** (input field) — optional subject keyword filter.
- **POLL_MS** (hardcoded) — polling interval in ms (30000).
- **LOOKBACK_MINUTES** (hardcoded) — startup lookback (60 min).
- **USE_TTS** (dropdown) — true => synthesize WAV via SAPI; false => attach static file.
- **VOICE_TEXT** (input field) — TTS text.
- **STATIC_AUDIO_PATH** (input field) — absolute path to MP3/WAV when USE_TTS=false.
- **REPLY_MODE** (dropdown) — "Reply" | "ReplyAll" | "NewToDL".
- **PROCESS_CATEGORY** (hardcoded) — category label to mark processed mails.
- **SKIP_SUBJECT_PATTERNS** and **SKIP_SENDER_PATTERNS** (hardcoded) — regex guards.

## Logic details
- **On load**:
  - Create COM instances, get Inbox (GetDefaultFolder(6)), sort Items by [ReceivedTime] desc.
  - Explore all available email accounts and automatically detect Gmail accounts.
  - Initialize UI with default values and event handlers.
  - Set lastCheck = now - LOOKBACK_MINUTES.
  - Wait for user to configure settings and start monitoring.

- **On each scan**:
  - Restrict("[ReceivedTime] >= 'MM/DD/YYYY HH:MM AM/PM'") (US-format).
  - For each IPM.Note:
    - Generate unique email identifier (EntryID or combo).
    - Skip if already processed by ID tracking or category.
    - Skip if sender/subject matches skip regex.
    - Check addressed to DL:
      - quick: mail.To/mail.CC display strings,
      - deeper: iterate Recipients → AddressEntry → GetExchangeUser().PrimarySmtpAddress.
    - Check subject filter (if specified).
    - Check keyword hit in Subject or Body.
    - If hit:
      - Add to processed list immediately (prevent duplicates).
      - Build reply according to REPLY_MODE.
      - Prepend canned HTML body.
      - Attach:
        - TTS WAV via SAPI (write to temp, attach, then delete),
        - or STATIC_AUDIO_PATH.
      - Send with comprehensive error handling.
      - Add PROCESS_CATEGORY to original mail and save; optionally mark read.
  - Advance lastCheck to newest processed time (+1s).

- **UI**:
  - Responsive window with configuration panel, status display, and log area.
  - Real-time status updates and monitoring controls.
  - Comprehensive logging with copy/clear functionality.
  - Multi-account support with manual selection capability.

- **Safety**:
  - Dual-layer deduplication prevents reply loops.
  - Skip auto-replies, noreply senders, self-sender.
  - Robust error handling with detailed logging.
  - JScript compatibility (no modern JS methods).

## Provide complete code for OutlookVoiceAgent.hta
- Include <HTA:APPLICATION> header with proper window settings.
- Inline CSS for modern, responsive UI with configuration panel.
- Full JS with:
  - Dynamic configuration system (no hardcoded values in logic).
  - Multi-account detection and selection.
  - Dual-layer deduplication system.
  - Comprehensive debugging and logging.
  - JScript-compatible helper functions (trimString, manual loops).
  - Enhanced error handling and recovery.
- Ensure *strong comments* for each logical block.

## Provide README.md
- **What it does / limitations**:
  - Needs Outlook desktop open.
  - HTA runs with high privileges; only run trusted code.
  - TTS produces *WAV*; if MP3 required, use USE_TTS=false with a static MP3.
  - Supports multiple email accounts with automatic Gmail detection.
  - Cannot process shared/team mailboxes unless accessible in current profile.

- **Quick start**:
  1) Launch OutlookVoiceAgent.hta (double-click or use start_agent.bat).
  2) Configure settings in the UI (DL, keywords, reply mode, voice options).
  3) Click "Start Monitoring" to begin; use "Scan Now" for immediate testing.
  4) Use "Copy Log" to export debug information if needed.

- **Multi-account setup**:
  - Application automatically detects Gmail accounts.
  - Use "Select Inbox" to manually choose from available accounts.
  - Use "Refresh Folders" to explore account structure.

- **Auto-start (Task Scheduler)** example:
  - Trigger: At log on
  - Action: mshta.exe with arguments "C:\Path\OutlookVoiceAgent.hta"
  - Optional: Run with highest privileges; Delay task 30s to let Outlook start.

- **Troubleshooting**:
  - **COM errors**: ensure Outlook desktop is running; verify Trust Center settings.
  - **No matches**: check DL string (use SMTP), verify account selection, check keywords.
  - **Duplicate replies**: Enhanced deduplication should prevent this; check logs.
  - **Account issues**: Use "Select Inbox" to choose correct email account.
  - **JavaScript errors**: Application uses JScript-compatible code only.

- **Debugging features**:
  - Comprehensive logging with timestamps and color coding.
  - "Copy Log" button to export debug information.
  - Detailed email processing step tracking.
  - Account and folder exploration capabilities.

- Security notes and how to uninstall.

## Provide start_agent.bat
- A one-liner to launch the agent: `start mshta.exe "%~dp0OutlookVoiceAgent.hta"`

## Key Improvements Made
1. **Dynamic UI Configuration**: No more code editing required.
2. **Multi-Account Support**: Automatic Gmail detection and manual selection.
3. **Dual-Layer Deduplication**: Prevents duplicate replies reliably.
4. **Comprehensive Debugging**: Detailed logging for troubleshooting.
5. **Enhanced Error Handling**: Graceful recovery from failures.
6. **JScript Compatibility**: Works with older JavaScript engines.
7. **User-Friendly Interface**: Modern UI with real-time status updates.
8. **Log Management**: Copy and clear log functionality.
9. **Robust Email Processing**: Step-by-step filtering with detailed feedback.
10. **Production Ready**: Clean, commented, maintainable code.