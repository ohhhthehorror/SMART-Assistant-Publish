# 📦 SMART Assistant - Release Guide

## Publishing New Versions to GitHub

### Step 1: Build the EXE

**Windows (Easy Method):**
1. Double-click `BUILD_EXE.bat`
2. Wait for build to complete
3. Find EXE in `dist/` folder

**Windows (PowerShell):**
```powershell
.\BUILD_EXE.ps1
```

**Manual Build:**
```bash
python build_exe.py
```

---

### Step 2: Test the EXE

**Before releasing, test:**
- ✅ App launches without errors
- ✅ Login works
- ✅ All features function correctly
- ✅ Settings save properly
- ✅ Auto-update check works

---

### Step 3: Create GitHub Release

1. **Go to your GitHub repository**
   - https://github.com/YOUR_USERNAME/smart-assistant

2. **Click "Releases"** (right sidebar)

3. **Click "Draft a new release"**

4. **Fill in release information:**
   ```
   Tag version: v0.2.1
   Release title: SMART Assistant v0.2.1
   
   Description:
   ## What's New
   - User deactivation with automated scanning
   - Batch deactivation from Excel/CSV
   - User standards compliance checking
   - Auto-update system
   - Dark/Light theme support
   
   ## Installation
   1. Download SMART_Assistant.exe
   2. Run the executable
   3. Follow first-time setup
   
   ## System Requirements
   - Windows 10/11
   - Internet connection (for SMART system access)
   - 250MB free disk space
   ```

5. **Upload the EXE:**
   - Drag `dist/SMART_Assistant.exe` into the "Attach binaries" area
   - File will upload automatically

6. **Click "Publish release"**

---

### Step 4: Update version.json

After publishing, update `version.json` in your repository:

```json
{
  "version": "0.2.1",
  "release_date": "2026-02-10",
  "download_url": "https://github.com/YOUR_USERNAME/smart-assistant/releases/download/v0.2.1/SMART_Assistant.exe",
  "release_notes": "User deactivation, batch operations, standards scanning, auto-updates",
  "minimum_version": "0.2.0",
  "changelog": [
    "User deactivation with scan functionality",
    "Batch deactivation from spreadsheet",
    "User standards compliance scanning",
    "Auto-update system",
    "Dark/Light theme support"
  ]
}
```

**Commit and push to GitHub:**
```bash
git add version.json
git commit -m "Update version.json for v0.2.1"
git push
```

---

### Step 5: Verify Auto-Update Works

1. Keep v0.2.1 EXE running
2. Publish a new version (v0.2.2) on GitHub
3. Update version.json with v0.2.2 info
4. Restart v0.2.1 app
5. Should see "Update Available" dialog
6. Click "Yes" to test auto-update

---

## Version Numbering

Follow semantic versioning: **MAJOR.MINOR.PATCH**

- **MAJOR** (1.0.0): Breaking changes, complete rewrites
- **MINOR** (0.2.0): New features, non-breaking changes
- **PATCH** (0.2.1): Bug fixes, small improvements

### Examples:
```
0.2.1 → 0.2.2  (Bug fix)
0.2.1 → 0.3.0  (New feature)
0.2.1 → 1.0.0  (Major rewrite)
```

---

## Quick Release Checklist

- [ ] Update version in `SMART_ASSISTANT_FINAL.py`
- [ ] Build EXE with `BUILD_EXE.bat`
- [ ] Test EXE thoroughly
- [ ] Create GitHub release
- [ ] Upload EXE to release
- [ ] Update `version.json` with new version
- [ ] Commit and push `version.json`
- [ ] Test auto-update from previous version

---

## File Structure in Repository

```
smart-assistant/
├── README.md
├── version.json          ← Update this after each release
├── releases/
│   ├── v0.2.1/
│   │   └── SMART_Assistant.exe
│   └── v0.2.2/
│       └── SMART_Assistant.exe
└── src/                  ← Optional: source code
    ├── SMART_ASSISTANT_FINAL.py
    ├── auto_updater.py
    ├── build_exe.py
    └── BUILD_EXE.bat
```

---

## Troubleshooting

### Build fails
```bash
# Clean everything and rebuild
rmdir /s /q build dist
del *.spec
BUILD_EXE.bat
```

### Auto-update not working
1. Check `version.json` is accessible
2. Verify GitHub release has .exe attached
3. Check download_url in version.json is correct
4. Test with: https://github.com/USER/REPO/releases/latest

### Users can't download
- Make sure GitHub release is **Published** (not Draft)
- Check .exe is attached to release
- Verify download_url is publicly accessible

---

## GitHub Repository Setup

### Initial Setup:
```bash
# Create repository on GitHub first, then:
git init
git add .
git commit -m "Initial commit - SMART Assistant v0.2.1"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/smart-assistant.git
git push -u origin main
```

### For Each Release:
```bash
# Update version in code
# Build EXE
# Create GitHub release
# Update version.json

git add version.json
git commit -m "Release v0.2.1"
git push
```

---

## Auto-Update Configuration

Edit in `SMART_ASSISTANT_FINAL.py`:

```python
# GitHub configuration for auto-updates
GITHUB_USER = "YOUR_USERNAME"      # ← Change this
GITHUB_REPO = "smart-assistant"    # ← Change this
```

**Current settings:**
- Checks for updates on startup
- User can skip updates
- Downloads to same folder as EXE
- Auto-replaces and restarts

---

## Support

If you encounter issues:
1. Check build output for errors
2. Verify Python 3.12+ is installed
3. Ensure all dependencies are installed
4. Try manual PyInstaller command
5. Check GitHub release is public

For auto-update issues:
1. Verify version.json is committed
2. Check GitHub release download URL
3. Test with Postman/browser
4. Review error messages in app
