# VBA MCP Server - Roadmap & TODO

## Current Status: v0.8.0 (Production Ready)

**Last Updated:** 2024-12-30

### Test Results
- Excel: 20/20 (100%)
- Access: 16/16 (100%)

---

## Implemented Features

### Excel Support (100%)
- [x] Extract VBA code from modules
- [x] Inject VBA code into modules
- [x] List all VBA modules
- [x] Run macros (Sub/Function)
- [x] Read worksheet data
- [x] Write worksheet data
- [x] Excel Tables (ListObjects) operations
- [x] VBA syntax validation
- [x] Backup/Rollback system

### Access Support (100%)
- [x] List tables with schema (fields, types, record count)
- [x] List saved queries (QueryDefs)
- [x] Execute SELECT queries
- [x] Execute INSERT queries
- [x] Execute UPDATE queries
- [x] Execute DELETE queries
- [x] Run saved queries
- [x] Read data with filters (WHERE, ORDER BY, LIMIT)
- [x] Write data (append mode)
- [x] Write data (replace mode)
- [x] Extract VBA via COM
- [x] Inject VBA modules
- [x] VBA syntax validation
- [x] Backup/Rollback system
- [x] Forms support (list, create, delete, export/import as text)

### Access VBA via COM (NEW in v0.8.0)
- [x] **extract_vba_access** - Extract VBA code from .accdb using COM
  - Uses VBProject.VBComponents instead of oletools
  - Full code extraction with procedure parsing
  - File: `packages/pro/src/vba_mcp_pro/tools/access_vba.py`

- [x] **analyze_structure_access** - Analyze VBA structure via COM
  - Complexity metrics (cyclomatic complexity)
  - Procedure analysis with call graph
  - Dependency detection
  - Recommendations for refactoring

- [x] **compile_vba** - Compile VBA and detect errors
  - Uses Application.DoCmd.RunCommand(acCmdCompileAndSaveAllModules)
  - Reports compilation errors with module/line
  - Detects best practice warnings (Option Explicit, etc.)
  - Works with Access and Excel files

---

## TODO - High Priority

### Security Enhancements
- [ ] **Parameterized queries for Access**
  - Prevent SQL injection attacks
  - Use `?` placeholders with parameter binding
  - Priority: HIGH
  - Difficulty: Medium

### Excel Enhancements
- [ ] **Pivot Tables support**
  - Create pivot tables from data ranges
  - Refresh existing pivot tables
  - Read pivot table data
  - Priority: HIGH
  - Difficulty: Hard

- [ ] **Named Ranges**
  - List named ranges
  - Create/Delete named ranges
  - Read/Write to named ranges
  - Priority: MEDIUM
  - Difficulty: Easy

- [ ] **Formulas support**
  - Read cell formulas (not just values)
  - Write formulas to cells
  - Evaluate formulas
  - Priority: MEDIUM
  - Difficulty: Easy

### Access Enhancements
- [ ] **Schema modification (DDL)**
  - CREATE TABLE
  - ALTER TABLE (add/modify/drop columns)
  - DROP TABLE
  - CREATE INDEX
  - Priority: MEDIUM
  - Difficulty: Easy

- [ ] **Transactions support**
  - BEGIN TRANSACTION
  - COMMIT
  - ROLLBACK
  - Priority: MEDIUM
  - Difficulty: Medium

- [ ] **Relationships in schema**
  - Expose foreign key relationships
  - Show table relationships
  - Priority: LOW
  - Difficulty: Easy

---

## TODO - Medium Priority

### Excel
- [ ] **Charts support**
  - Create basic charts (bar, line, pie)
  - Modify existing charts
  - Export charts as images
  - Priority: MEDIUM
  - Difficulty: Medium

- [ ] **Conditional Formatting**
  - Apply conditional formatting rules
  - Read existing rules
  - Priority: LOW
  - Difficulty: Medium

- [ ] **Data Validation**
  - Set cell validation rules
  - Dropdown lists
  - Priority: LOW
  - Difficulty: Medium

### Access
- [x] ~~**Forms support**~~ → Implemented in v0.7.0
  - list_access_forms, create_access_form, delete_access_form
  - export_form_definition (SaveAsText), import_form_definition (LoadFromText)

- [ ] **Reports support** (read-only)
  - List reports
  - Export reports to PDF
  - Priority: LOW
  - Difficulty: Hard

- [ ] **Linked Tables**
  - List linked tables
  - Refresh links
  - Priority: LOW
  - Difficulty: Medium

---

## TODO - Low Priority / Future

### New Office Applications
- [ ] **Word support**
  - Extract VBA from .docm
  - Inject VBA
  - Run macros
  - Read/Write document content
  - Priority: LOW
  - Difficulty: Medium

- [ ] **PowerPoint support**
  - Extract VBA from .pptm
  - Inject VBA
  - Run macros
  - Priority: LOW
  - Difficulty: Medium

### Advanced Features
- [ ] **Batch operations**
  - Process multiple files
  - Bulk VBA injection
  - Priority: LOW
  - Difficulty: Medium

- [ ] **Event hooks**
  - Before/After save events
  - Macro execution events
  - Priority: LOW
  - Difficulty: Hard

- [ ] **Performance optimizations**
  - Connection pooling
  - Lazy loading
  - Caching
  - Priority: LOW
  - Difficulty: Medium

---

## Known Limitations

### Access
1. **oletools** - Doesn't support .accdb files
   - Affected tools: `extract_vba`, `list_modules`, `analyze_structure` (lite version)
   - SOLVED in v0.8.0: Use `extract_vba_access`, `analyze_structure_access` instead
   - These Pro tools use COM (VBProject.VBComponents) instead of oletools

2. **run_macro** - Application.Run behaves differently than Excel
   - Access macros (UI-created) vs VBA procedures are different
   - Must use just procedure name, not "Module.Procedure" format
   - If VBA project has compilation errors, run_macro fails silently
   - SOLVED in v0.8.0: Use `compile_vba` to detect errors first

3. **DisplayAlerts** - Access doesn't have this property
   - Non-blocking, just a warning message

### Excel
1. **Rapid injections** - 5+ injections in <2s may crash COM
   - Workaround: Add small delay between operations

### General
1. **Windows only** - Requires pywin32 + COM
2. **Trust VBA** - Must enable "Trust access to VBA project object model"

---

## Contributing

To contribute:
1. Pick an item from TODO
2. Create a feature branch
3. Implement with tests
4. Submit PR

Priority order: HIGH > MEDIUM > LOW

---

## Version History

| Version | Date | Highlights |
|---------|------|------------|
| v0.8.0 | 2024-12-30 | Access VBA via COM (extract, analyze, compile) |
| v0.7.0 | 2024-12-30 | Access Forms support (5 tools) |
| v0.6.0 | 2024-12-30 | Complete Access support, action queries |
| v0.5.0 | 2024-12-28 | VBA injection fixes, validation improvements |
| v0.4.0 | 2024-12-15 | Excel Tables, critical fixes |
| v0.3.0 | 2024-12-10 | Initial Access support |
| v0.2.0 | 2024-12-05 | Session manager, backup system |
| v0.1.0 | 2024-12-01 | Initial release, Excel VBA support |
