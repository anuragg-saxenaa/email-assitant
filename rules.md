## Goal
Create an advanced HTA application that:
- **Smart Email Monitoring**: Polls Outlook Inbox with user-configurable intervals (10s-5min).
- **Flexible Time Windows**: User-configurable lookback windows (15min-24hrs) for email detection.
- **Multi-Account Support**: Automatically detects Gmail accounts with manual selection capability.
- **Advanced Filtering**: Filters messages addressed to specific *Distribution Lists (DL)* with intelligent matching.
- **Comprehensive Processing**: Scans *subject/body* for configured *keywords* with detailed logging.
- **Automated Responses**: If matched, *auto-replies* (Reply/ReplyAll/NewToDL) with:
  - *Predefined voice notes*: TTS-generated WAV or static MP3/WAV files
  - Professional HTML message bodies with customizable content
- **Bulletproof Deduplication**: Dual-layer system using Category tags + Email ID tracking.
- **Smart Filtering**: Automatically skips auto-replies, OOO messages, and noreply senders.
- **User-Friendly Interface**: Modern UI with comprehensive configuration and real-time diagnostics.
- **Enhanced Logging**: Detailed processing summaries with actionable troubleshooting suggestions.

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

### **Multi-Account Support**
- **Automatic Detection**: Scans all Outlook stores and identifies Gmail accounts
- **Manual Selection**: "Select Inbox" button for choosing specific email accounts
- **Account Exploration**: "Refresh Folders" for detailed account structure analysis
- **Smart Switching**: Automatically prefers Gmail accounts when detected

### **Bulletproof Deduplication System**
- **Layer 1**: Outlook category-based tracking ("AI-VoiceReplied")
- **Layer 2**: Email ID tracking using EntryID or sender+subject+timestamp combo
- **Layer 3**: Reply-sent tracking to prevent race conditions
- **Layer 4**: **Reply Loop Prevention** - Detects and skips own automated replies
- **Persistent Memory**: Maintains processed email lists across sessions (up to 200 entries)
- **Early Marking**: Marks emails as processed BEFORE sending to prevent duplicates

### **Comprehensive Diagnostic System**
- **Step-by-Step Processing**: Detailed logs for every email processing decision
- **Email Processing Summaries**: Post-scan reports with skip reasons and suggestions
- **Time Filtering Debug**: Detailed time comparison and filter testing
- **Configuration Validation**: Real-time validation with specific error messages
- **Actionable Suggestions**: Specific guidance for fixing configuration issues

### **Advanced Error Handling**
- **Graceful Degradation**: Continue operation even when individual components fail
- **Multiple Fallbacks**: Alternative methods for clipboard, file operations, etc.
- **Context-Rich Logging**: Detailed error messages with troubleshooting context
- **Recovery Mechanisms**: Automatic retry and alternative approaches
- **HTA Compatibility**: JScript-compatible code for maximum reliability

## Config (Dynamic UI Inputs)
All parameters now configurable through modern UI without code editing:

### **User-Configurable Settings**
- **DL_ADDRESS** (input field) — Distribution List SMTP address or display name to match
- **KEYWORDS** (textarea) — Comma-separated, case-insensitive subject/body triggers
- **SUBJECT_FILTER** (input field) — Optional additional subject keyword filter
- **LOOKBACK_WINDOW** (dropdown) — Startup lookback window:
  - Options: 15min, 30min, 1hr (default), 2hr, 4hr, 8hr, 24hr
- **POLL_FREQUENCY** (dropdown) — Email checking interval:
  - Options: 10sec, 15sec, 30sec (default), 1min, 2min, 5min
- **USE_TTS** (dropdown) — Voice mode: TTS generation or static audio file
- **VOICE_TEXT** (input field) — Custom text for TTS synthesis
- **STATIC_AUDIO_PATH** (input field) — Absolute path to MP3/WAV when using static files
- **REPLY_MODE** (dropdown) — Response method: "Reply" | "ReplyAll" | "NewToDL"

### **System Settings (Hardcoded)**
- **PROCESS_CATEGORY** — Category label "AI-VoiceReplied" to mark processed emails
- **SKIP_SUBJECT_PATTERNS** — Auto-skip patterns for OOO, auto-replies, etc.
- **SKIP_SENDER_PATTERNS** — Auto-skip patterns for noreply addresses

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
    - **Multi-layer duplicate prevention**:
      - Skip if already processed by ID tracking.
      - Skip if reply already sent to this email.
      - Skip if already processed by category.
    - **Smart filtering**:
      - Skip if sender/subject matches skip patterns.
      - **Skip if own automated reply** (prevent reply loops):
        - Check body for "This is an automated voice response".
        - Check for multiple "RE:" in subject line.
        - Check for AI-VoiceReplied category.
    - **Address and keyword validation**:
      - Check addressed to DL (display strings + SMTP resolution).
      - Check subject filter (if specified).
      - Check keyword hit in Subject or Body.
    - **If all checks pass**:
      - Mark as processed immediately (prevent race conditions).
      - Mark reply as sent (prevent duplicate replies).
      - Build reply according to REPLY_MODE.
      - Attach voice note (TTS WAV or static audio).
      - Send with comprehensive error handling.
      - Add PROCESS_CATEGORY and save.
  - **Enhanced logging**: Generate processing summary with skip reasons and suggestions.
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

## Key Improvements Made (Version 2.0)

### **🎛️ User Experience Enhancements**
1. **Dynamic UI Configuration**: Complete elimination of code editing requirements
2. **Flexible Timing Controls**: User-configurable lookback windows and poll frequencies
3. **Modern Interface**: Responsive design with real-time status updates and progress indicators
4. **Enhanced Log Management**: HTA-compatible copy functionality with multiple fallback methods
5. **Smart Diagnostics**: Comprehensive processing summaries with actionable troubleshooting guidance

### **🔧 Technical Improvements**
6. **Multi-Account Architecture**: Automatic Gmail detection with manual selection capabilities
7. **Triple-Layer Deduplication**: Category + ID + Reply tracking for bulletproof duplicate prevention
8. **Advanced Time Handling**: Sophisticated date formatting and time window management
9. **Robust Error Recovery**: Multiple fallback mechanisms with graceful degradation
10. **JScript Optimization**: Full compatibility with older JavaScript engines in HTA environment

### **📊 Monitoring & Debugging**
11. **Comprehensive Logging System**: Step-by-step processing with detailed decision tracking
12. **Email Processing Summaries**: Post-scan analysis with specific skip reasons and fix suggestions
13. **Configuration Validation**: Real-time validation with context-specific error messages
14. **Debug Mode Enhancements**: Detailed time comparisons and filter testing capabilities
15. **Production Readiness**: Clean, well-commented, maintainable codebase with extensive error handling

### **🚀 Performance & Reliability**
16. **Memory Management**: Efficient tracking lists with automatic cleanup (200-item limits)
17. **Race Condition Prevention**: Early marking system prevents duplicate processing
18. **Reply Loop Prevention**: Multi-method detection of own automated replies
19. **Persistent State Management**: Maintains processing history across application restarts
20. **Smart Resource Handling**: Automatic cleanup of temporary files and COM objects
21. **Scalable Architecture**: Designed to handle high-volume email environments efficiently

### **🔄 Reply Loop Prevention (Latest Addition)**
22. **Body Content Detection**: Identifies replies containing "This is an automated voice response"
23. **Subject Pattern Analysis**: Detects multiple "RE:" prefixes indicating reply chains
24. **Category Cross-Check**: Verifies against AI-VoiceReplied category as backup
25. **Comprehensive Logging**: Detailed reasons for skipping automated replies with troubleshooting info