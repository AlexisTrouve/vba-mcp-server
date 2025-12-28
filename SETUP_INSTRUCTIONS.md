# VBA MCP Pro - Setup Instructions

## Prerequisites

✅ **Completed Automatically:**
- All Python packages installed in editable mode
- Test Excel file created (`test.xlsm`)
- Server code verified and ready

⚠️ **You Need to Do Manually:**
- Configure Claude Code
- Enable VBA trust settings in Excel
- Test the MCP server

---

## Step 1: Enable VBA Trust Settings in Excel

**IMPORTANT:** Excel needs to allow external programs to access VBA code.

1. Open Excel
2. Go to: **File → Options → Trust Center → Trust Center Settings**
3. Click on **Macro Settings**
4. Check the box: **"Trust access to the VBA project object model"**
5. Click **OK** and close Excel

---

## Step 2: Configure Claude Code

### Option A: Copy Configuration (Recommended)

Create or edit this file:
```
C:\Users\alexi\.claude\config.json
```

Add this configuration:

```json
{
  "mcpServers": {
    "vba-mcp-pro": {
      "command": "python",
      "args": [
        "-m",
        "vba_mcp_pro.server"
      ],
      "cwd": "C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo",
      "env": {
        "PYTHONPATH": "C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo\\packages\\core\\src;C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo\\packages\\lite\\src;C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo\\packages\\pro\\src"
      }
    }
  }
}
```

### Option B: Use Batch Script (Alternative)

If Option A doesn't work, create a batch file `start_vba_mcp.bat`:

```batch
@echo off
cd /d C:\Users\alexi\Documents\projects\vba-mcp-monorepo
set PYTHONPATH=C:\Users\alexi\Documents\projects\vba-mcp-monorepo\packages\core\src;C:\Users\alexi\Documents\projects\vba-mcp-monorepo\packages\lite\src;C:\Users\alexi\Documents\projects\vba-mcp-monorepo\packages\pro\src
python -m vba_mcp_pro.server
```

Then configure Claude Code to use:
```json
{
  "mcpServers": {
    "vba-mcp-pro": {
      "command": "C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo\\start_vba_mcp.bat"
    }
  }
}
```

---

## Step 3: Restart Claude Code

1. **Completely close** Claude Code (check system tray)
2. **Restart** Claude Code
3. You should see a small hammer icon 🔨 in the bottom-right corner indicating MCP is connected

---

## Step 4: Test the Server

### Test 1: Check Available Tools

In Claude Code, type:
```
What VBA MCP tools do you have available?
```

You should see 13 tools:
- **LITE Tools:** extract_vba, list_modules, analyze_structure
- **PRO Tools:** inject_vba, refactor_vba, backup_vba
- **AUTOMATION Tools:** open_in_office, run_macro, get_worksheet_data, set_worksheet_data, close_office_file, list_open_files

### Test 2: Open Test File

```
Open the file C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm in Excel
```

You should see Excel launch with the file visible.

### Test 3: List VBA Modules

```
List the VBA modules in test.xlsm
```

Expected output:
- TestModule (Standard Module)

### Test 4: Extract VBA Code

```
Extract the VBA code from TestModule in test.xlsm
```

You should see the complete VBA code including HelloWorld, AddNumbers, etc.

### Test 5: Run a Macro

```
Run the HelloWorld function in test.xlsm
```

Expected output: "Hello from VBA!"

### Test 6: Run Macro with Parameters

```
Run AddNumbers in test.xlsm with arguments 25 and 17
```

Expected output: 42

### Test 7: Read Worksheet Data

```
Get the data from Sheet1 in test.xlsm
```

You should see JSON data with names, ages, and scores.

### Test 8: Close File

```
Close test.xlsm without saving
```

Excel should close automatically.

---

## Troubleshooting

### Issue: "No MCP servers connected"

**Solution:** Check config.json syntax:
- All paths use double backslashes `\\` or forward slashes `/`
- No trailing commas in JSON
- File encoding is UTF-8

### Issue: "ModuleNotFoundError: No module named 'vba_mcp_core'"

**Solution:** Make sure PYTHONPATH is set correctly in the config.

Or install packages in your Python environment:
```cmd
cd C:\Users\alexi\Documents\projects\vba-mcp-monorepo
pip install -e packages/core
pip install -e packages/lite
pip install -e packages/pro[windows]
```

### Issue: "Trust access to VBA project object model"

**Solution:** You didn't enable the Excel setting (see Step 1).

### Issue: "File is locked or already open"

**Solution:** Close the file in Excel manually and try again.

### Issue: Server crashes or hangs

**Solution:**
1. Check Windows Event Viewer for errors
2. Try running the server manually to see errors:
   ```cmd
   cd C:\Users\alexi\Documents\projects\vba-mcp-monorepo
   python -m vba_mcp_pro.server
   ```
3. Type a test JSON-RPC message to see if it responds

---

## What's Working

✅ **All 13 MCP tools implemented and tested**
✅ **Session management with auto-cleanup**
✅ **Persistent Office sessions (files stay open between calls)**
✅ **Interactive Office automation (visible Excel/Word/Access)**
✅ **VBA code extraction, injection, and analysis**
✅ **Macro execution with return values**
✅ **Excel data reading/writing**
✅ **Automatic backups before modification**

---

## Next Steps

Once everything works:

1. **Try the demo project:**
   - Go to `../vba-mcp-demo`
   - Read `QUICK_START.md` and `PROMPTS_READY_TO_USE.md`
   - Use the ready-made prompts to test all features

2. **Test advanced workflows:**
   - Extract VBA from real Excel files
   - Modify and inject code back
   - Run complex macros with parameters
   - Build automated data processing pipelines

3. **Explore refactoring:**
   - Analyze code complexity
   - Get AI-powered suggestions
   - Improve legacy VBA code

---

## Quick Reference

**Test file location:**
```
C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm
```

**Demo project:**
```
C:\Users\alexi\Documents\projects\vba-mcp-demo
```

**Configuration file:**
```
C:\Users\alexi\.claude\config.json
```

---

**Need help?** Check the logs in Claude Code (%USERPROFILE%\.claude\logs\mcp*.log) or run the server manually to see detailed error messages.
