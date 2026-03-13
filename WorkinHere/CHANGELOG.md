# SMART Assistant - Changelog

## v0.2.10 (2026-02-11)

### 🐛 Bug Fixes
- Fixed auto-update not launching new version after download
- Fixed Chrome window close causing connection error crash
- Fixed popup text cutoff issues (all dialogs now dynamically sized)

### ✨ Improvements
- Moved version number to top left corner (was bottom center)
- Update dialog now matches app theme (dark/light mode)
- Update check now happens before login screen
- Added "Stop Scan" button to User Standards Scan
- Improved error handling in auto-updater with fallback dialogs
- User Standards Scan now uses exact same automation as deactivation scan

### 🔧 Technical
- Updated batch script with better directory handling (`cd /d` and `start /D`)
- Added comprehensive error logging to updater
- Simplified update callback logic
- Better Chrome connection handling in scans
- Fixed page size expansion (Shift+Tab 14x navigation)
- Improved sort reliability

---

## v0.2.1 (2026-02-10)

### 🎉 Initial Release

### Features
- **User Deactivation**
  - Automated scanning with Chrome automation
  - Batch deactivation from Excel/CSV spreadsheets
  - Inline editing in spreadsheet view
  - Ticket number tracking
  
- **User Standards Scan**
  - Username format validation (xxxxflast)
  - Job title compliance checking
  - Real-time statistics display
  - Detailed console logging

- **Equipment Standards Scan**
  - Placeholder for future implementation

- **Auto-Update System**
  - GitHub-based release checking
  - Automatic download and installation
  - Version tracking with version.json

- **User Interface**
  - Dark/Light theme toggle
  - Headless/Visible Chrome mode
  - Automation speed control
  - Settings persistence

- **Build System**
  - One-click EXE builder (BUILD_EXE.bat)
  - PyInstaller integration
  - Icon support
  - Auto-bundling of dependencies

---

## Version Numbering

We use semantic versioning: **MAJOR.MINOR.PATCH**

- **MAJOR** (1.0.0): Breaking changes, major rewrites
- **MINOR** (0.2.0): New features, non-breaking additions
- **PATCH** (0.2.1): Bug fixes, small improvements

### Examples:
- Bug fix: 0.2.1 → 0.2.2
- New feature: 0.2.2 → 0.3.0
- Major rewrite: 0.2.2 → 1.0.0
