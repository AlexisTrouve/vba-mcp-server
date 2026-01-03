"""Pro-only MCP tools for VBA manipulation."""

from .inject import inject_vba_tool
from .refactor import refactor_tool
from .backup import backup_tool
from .validate import validate_vba_code_tool
from .office_automation import (
    open_in_office_tool,
    run_macro_tool,
    get_worksheet_data_tool,
    set_worksheet_data_tool,
    close_office_file_tool,
    list_open_files_tool,
    list_macros_tool,
    list_access_tables_tool,
    list_access_queries_tool,
    run_access_query_tool,
    # Access Forms tools
    list_access_forms_tool,
    create_access_form_tool,
    delete_access_form_tool,
    export_form_definition_tool,
    import_form_definition_tool
)
from .excel_tables import (
    list_tables_tool,
    insert_rows_tool,
    delete_rows_tool,
    insert_columns_tool,
    delete_columns_tool,
    create_table_tool
)
from .access_vba import (
    extract_vba_access_tool,
    analyze_structure_access_tool,
    compile_vba_tool
)

__all__ = [
    "inject_vba_tool",
    "refactor_tool",
    "backup_tool",
    "validate_vba_code_tool",
    "open_in_office_tool",
    "run_macro_tool",
    "get_worksheet_data_tool",
    "set_worksheet_data_tool",
    "close_office_file_tool",
    "list_open_files_tool",
    "list_macros_tool",
    "list_access_tables_tool",
    "list_access_queries_tool",
    "run_access_query_tool",
    # Access Forms tools
    "list_access_forms_tool",
    "create_access_form_tool",
    "delete_access_form_tool",
    "export_form_definition_tool",
    "import_form_definition_tool",
    # Excel Tables tools
    "list_tables_tool",
    "insert_rows_tool",
    "delete_rows_tool",
    "insert_columns_tool",
    "delete_columns_tool",
    "create_table_tool",
    # Access VBA tools (COM-based)
    "extract_vba_access_tool",
    "analyze_structure_access_tool",
    "compile_vba_tool"
]
