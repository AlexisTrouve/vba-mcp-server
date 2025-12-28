# 🚀 START HERE - VBA MCP Pro Setup

Everything has been prepared automatically. You just need to complete 3 manual steps.

---

## ✅ What's Already Done (Automatic)

- ✅ All Python packages installed (`core`, `lite`, `pro`)
- ✅ All dependencies installed (`pywin32`, `oletools`, etc.)
- ✅ Test Excel file created (`test.xlsm`) with 6 VBA functions
- ✅ **21 MCP tools** implemented and ready (v0.3.0):
  - 3 LITE tools (read-only)
  - 3 PRO tools (injection/validation)
  - 6 AUTOMATION tools (interactive Office)
  - 6 EXCEL TABLES tools (NEW v0.3.0!)
  - 2 BACKUP/REFACTOR tools
  - 1 LIST MACROS tool
- ✅ Session manager with auto-cleanup
- ✅ Configuration files created
- ✅ Documentation and test prompts ready

---

## ⚠️ What YOU Need to Do (3 Steps - 5 Minutes)

### Step 1: Enable Excel VBA Trust (2 minutes)

1. **Open Excel**
2. **File → Options → Trust Center → Trust Center Settings**
3. **Macro Settings**
4. **Check:** "Trust access to the VBA project object model"
5. **Click OK** and close Excel

**Why?** This allows Python to read/modify VBA code in Excel files.

---

### Step 2: Configure Claude Code (2 minutes)

#### Option A: Copy the Configuration File (Easiest)

1. **Copy this file:**
   ```
   C:\Users\alexi\Documents\projects\vba-mcp-monorepo\claude_desktop_config.json
   ```
   (Note: This is an example config file for reference)

2. **To this location:**
   ```
   %USERPROFILE%\.claude\config.json
   ```

   **If the file already exists:** Open it and merge the content (add the "vba-mcp-pro" entry to the "mcpServers" section)

#### Option B: Manual Edit (Alternative)

1. Open (or create):
   ```
   %USERPROFILE%\.claude\config.json
   ```

2. Paste this content:
   ```json
   {
     "mcpServers": {
       "vba-mcp-pro": {
         "command": "python",
         "args": ["-m", "vba_mcp_pro.server"],
         "cwd": "C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo",
         "env": {
           "PYTHONPATH": "C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo\\packages\\core\\src;C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo\\packages\\lite\\src;C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo\\packages\\pro\\src"
         }
       }
     }
   }
   ```

---

### Step 3: Test in Claude Code (1 minute)

1. **Completely close** Claude Code (check system tray!)
2. **Restart** Claude Code
3. **Look for:** Small hammer icon 🔨 in bottom-right (means MCP connected)
4. **Type this prompt:**
   ```
   What VBA MCP tools do you have available?
   ```

**Expected result:** Claude lists **21 tools** including `open_in_office`, `run_macro`, `extract_vba`, `list_tables`, `create_table`, etc.

---

## 🎯 Quick Test (30 seconds)

Once Claude Code shows the tools, try this:

```
Open C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm in Excel
```

**Expected:** Excel launches with the test file visible.

Then:
```
Run the HelloWorld function in test.xlsm
```

**Expected:** Claude returns "Hello from VBA!"

### Test Excel Tables (NEW v0.3.0) - 30 seconds

```
In test.xlsm:
1. Create an Excel table named "TestData" from range A1:C10 on Sheet1
2. List all tables to confirm it was created
```

**Expected:** Claude creates the table and confirms it exists with dimensions and headers.

---

## 📚 What to Do Next

### If Everything Works ✅

**→ Go to:** `QUICK_TEST_PROMPTS.md`
- Copy-paste 13 ready-to-use test prompts
- Test all features in 5 minutes

**→ Then go to:** `../vba-mcp-demo/PROMPTS_READY_TO_USE.md`
- 50+ advanced prompts for real-world scenarios

### If Something Doesn't Work ❌

**→ Go to:** `SETUP_INSTRUCTIONS.md`
- Detailed troubleshooting guide
- Alternative configuration methods
- Common error solutions

---

## 📂 File Reference

| File | Purpose |
|------|---------|
| `START_HERE.md` | **← You are here** - Quick start guide |
| `SETUP_INSTRUCTIONS.md` | Detailed setup & troubleshooting |
| `QUICK_TEST_PROMPTS.md` | 13 copy-paste test prompts |
| `claude_desktop_config.json` | MCP configuration example (for reference) |
| `start_vba_mcp.bat` | Alternative launcher script |
| `test.xlsm` | Test Excel file with VBA macros |
| `../vba-mcp-demo/` | Complete demo project with examples |

---

## 🔥 Feature Highlights

### What Can You Do with VBA MCP Pro?

1. **Extract VBA code** from any Office file (read-only)
2. **Modify VBA code** using Claude's AI
3. **Inject code back** into files (with automatic backup)
4. **Open Excel/Word/Access** visibly on screen
5. **Run VBA macros** with parameters and return values
6. **Read/Write Excel data** as JSON arrays
7. **Excel Tables (NEW v0.3.0)** - Structured table operations:
   - Create tables from ranges with formatting
   - Insert/delete rows and columns
   - Read/write by column names
   - List all tables with metadata
8. **Validate VBA code** before injection (syntax checking)
9. **Analyze code complexity** and get refactoring suggestions
10. **Automate workflows** combining multiple operations

### Interactive Sessions

Files stay open between operations (1-hour timeout):
```
1. Open file          → Excel launches
2. Run macro         → Uses same session (fast!)
3. Get data          → Still same session
4. Modify & inject   → Same session
5. Close file        → Excel closes
```

---

## ⚡ Pro Tips

- **Files auto-close** after 1 hour of inactivity (prevents memory leaks)
- **Automatic backups** are created before any VBA injection
- **Multiple files** can be open simultaneously
- **Claude Code logs** are helpful for debugging (%USERPROFILE%\.claude\logs\mcp*.log)

---

## 🎓 Learning Path

1. **Start:** Test basic extraction with `test.xlsm`
2. **Practice:** Use prompts from `QUICK_TEST_PROMPTS.md`
3. **Explore:** Try the demo project in `../vba-mcp-demo/`
4. **Apply:** Use with your real Excel files!

---

## 📞 Need Help?

**Check logs:**
- Claude Code: %USERPROFILE%\.claude\logs\mcp*.log
- Or run manually: `start_vba_mcp.bat` (see errors in console)

**Common issues:**
- "No MCP servers connected" → Check JSON syntax in config.json
- "Cannot run macro" → Enable Excel Trust Settings (Step 1)
- "File is locked" → Close file in Excel manually
- "Module not found" → Check PYTHONPATH in config.json

---

**Ready?** Complete the 3 steps above and start testing! 🚀

**Total time:** ~5 minutes for setup + testing
