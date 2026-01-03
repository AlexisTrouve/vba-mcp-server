# Plan: Support Formulaires Access

**Date:** 2025-12-30
**Version cible:** v0.7.0
**Status:** ✅ TERMINÉ

---

## Objectif

Ajouter la gestion complète des formulaires Access via 5 outils MCP, en exploitant `SaveAsText`/`LoadFromText` pour permettre à Claude de manipuler les formulaires comme du texte.

## Philosophie

```
┌─────────────────────────────────────────────────────────┐
│  100% des possibilités avec ~40% d'effort               │
│                                                         │
│  • Claude manipule du texte = son point fort            │
│  • Access exporte/importe en texte                      │
│  • Pas de limites artificielles                         │
└─────────────────────────────────────────────────────────┘
```

---

## Outils à Implémenter (5)

| Outil | Fonction | Priorité |
|-------|----------|----------|
| `list_access_forms` | Lister tous les formulaires | P0 |
| `create_access_form` | Créer formulaire (vide ou auto-généré) | P0 |
| `delete_access_form` | Supprimer un formulaire | P1 |
| `export_form_definition` | SaveAsText → fichier .txt | P0 |
| `import_form_definition` | LoadFromText ← fichier .txt | P0 |

---

## Phase 1: Recherche Format SaveAsText

### 1.1 Explorer le format texte Access

**Objectif:** Comprendre la structure du fichier exporté par SaveAsText

**Actions:**
1. Créer un formulaire simple dans Access (quelques contrôles)
2. Exporter avec `Application.SaveAsText acForm, "FormName", "path.txt"`
3. Analyser la structure du fichier

**Format attendu (approximatif):**
```
Version =21
VersionRequired =20
Begin Form
    RecordSource ="TableName"
    Caption ="Form Title"
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =6
    GridY =6
    Width =8000
    DatasheetFontHeight =11
    ...
    Begin
        Begin Label
            OverlapFlags =85
            Left =500
            Top =300
            Width =1500
            Height =300
            Name ="Label1"
            Caption ="Nom:"
        End
        Begin TextBox
            OverlapFlags =85
            Left =2000
            Top =300
            Width =3000
            Height =300
            Name ="txtNom"
            ControlSource ="Nom"
        End
        Begin CommandButton
            OverlapFlags =85
            Left =500
            Top =1000
            Width =2000
            Height =400
            Name ="btnSave"
            Caption ="Enregistrer"
            OnClick ="[Event Procedure]"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSave_Click()
    DoCmd.Save
End Sub
```

**Livrables:**
- [ ] Document de référence du format SaveAsText
- [ ] Exemples de différents types de contrôles
- [ ] Mapping propriétés → valeurs

---

### 1.2 Tester SaveAsText/LoadFromText manuellement

**Script VBA de test:**
```vba
' Dans Access, exécuter:
Sub TestSaveLoadForm()
    ' Exporter
    Application.SaveAsText acForm, "frm_Test", "C:\temp\frm_Test.txt"

    ' Modifier le fichier texte manuellement...

    ' Supprimer l'original
    DoCmd.DeleteObject acForm, "frm_Test"

    ' Réimporter
    Application.LoadFromText acForm, "frm_Test_Modified", "C:\temp\frm_Test.txt"
End Sub
```

**Points à valider:**
- [ ] Export fonctionne pour tous types de contrôles
- [ ] Import recrée le formulaire identique
- [ ] Le code VBA behind est inclus
- [ ] Les propriétés d'événements sont préservées

---

## Phase 2: Implémentation des Outils

### 2.1 list_access_forms

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/office_automation.py`

```python
async def list_access_forms_tool(file_path: str) -> str:
    """
    [PRO] List all forms in an Access database.

    Args:
        file_path: Path to .accdb file

    Returns:
        Formatted list of forms with metadata
    """
    session = await get_or_create_session(file_path)

    if session.app_type != "Access":
        raise ValueError("list_access_forms only works with Access files")

    app = session.app
    db = app.CurrentDb()

    forms = []

    # Méthode 1: Via AllForms collection
    for form in app.CurrentProject.AllForms:
        forms.append({
            "name": form.Name,
            "is_loaded": form.IsLoaded,
            "type": "Form"
        })

    # Formatage sortie
    output = "## Access Forms\n\n"
    output += f"**Database:** {Path(file_path).name}\n"
    output += f"**Total forms:** {len(forms)}\n\n"

    if forms:
        output += "| Name | Loaded |\n|------|--------|\n"
        for f in forms:
            loaded = "Yes" if f["is_loaded"] else "No"
            output += f"| {f['name']} | {loaded} |\n"
    else:
        output += "_No forms found_\n"

    return output
```

**Tests:**
```python
async def test_list_access_forms():
    result = await list_access_forms_tool(ACCESS_FILE)
    assert "frm_Employees" in result or "No forms found" in result
```

---

### 2.2 create_access_form

**Deux modes:**
1. **Vide** - Formulaire blank
2. **Auto-généré** - Basé sur une table/requête source

```python
async def create_access_form_tool(
    file_path: str,
    form_name: str,
    record_source: Optional[str] = None,
    auto_generate: bool = False,
    form_type: str = "single"  # "single", "continuous", "datasheet"
) -> str:
    """
    [PRO] Create a new Access form.

    Args:
        file_path: Path to .accdb file
        form_name: Name for the new form
        record_source: Table or query name to bind to
        auto_generate: If True, auto-create controls for all fields
        form_type: "single" (default), "continuous", or "datasheet"

    Returns:
        Success message with form details

    Examples:
        # Empty form
        create_access_form("db.accdb", "frm_New")

        # Form bound to table
        create_access_form("db.accdb", "frm_Clients",
                          record_source="Clients")

        # Auto-generated CRUD form
        create_access_form("db.accdb", "frm_Clients",
                          record_source="Clients",
                          auto_generate=True)
    """
    session = await get_or_create_session(file_path)
    app = session.app

    # Vérifier que le form n'existe pas déjà
    for form in app.CurrentProject.AllForms:
        if form.Name.lower() == form_name.lower():
            raise ValueError(f"Form '{form_name}' already exists")

    # Créer le formulaire
    frm = app.CreateForm()

    # Définir les propriétés de base
    if record_source:
        frm.RecordSource = record_source

    # Type de vue
    view_map = {"single": 0, "continuous": 1, "datasheet": 2}
    frm.DefaultView = view_map.get(form_type, 0)

    # Auto-génération des contrôles
    if auto_generate and record_source:
        await _auto_generate_form_controls(app, frm, record_source)

    # Sauvegarder avec le nom voulu
    app.DoCmd.Save(0, form_name)  # acForm = 0
    app.DoCmd.Close(0, frm.Name)

    # Renommer si nécessaire (CreateForm génère un nom temporaire)
    if frm.Name != form_name:
        app.DoCmd.Rename(form_name, 0, frm.Name)

    return f"""## Form Created

**Name:** {form_name}
**Record Source:** {record_source or "None (unbound)"}
**Type:** {form_type}
**Auto-generated:** {"Yes" if auto_generate else "No"}

Form created successfully. Use `export_form_definition` to view/edit the form structure.
"""


async def _auto_generate_form_controls(app, frm, record_source: str):
    """Generate TextBox + Label for each field in the record source."""
    db = app.CurrentDb()

    # Obtenir les champs de la source
    try:
        rs = db.OpenRecordset(record_source)
    except Exception:
        # Peut-être une requête
        rs = db.OpenRecordset(f"SELECT * FROM [{record_source}]")

    y_position = 500  # Position Y de départ
    y_spacing = 400   # Espacement entre contrôles

    for i in range(rs.Fields.Count):
        field = rs.Fields(i)
        field_name = field.Name

        # Label
        lbl = app.CreateControl(
            frm.Name,
            100,  # acLabel
            0,    # acDetail section
            "",   # Parent
            "",   # ColumnName
            200,  # Left
            y_position,  # Top
            1500, # Width
            300   # Height
        )
        lbl.Caption = f"{field_name}:"

        # TextBox
        txt = app.CreateControl(
            frm.Name,
            109,  # acTextBox
            0,    # acDetail section
            "",   # Parent
            field_name,  # ControlSource
            2000, # Left
            y_position,  # Top
            3000, # Width
            300   # Height
        )
        txt.Name = f"txt{field_name}"

        y_position += y_spacing

    rs.Close()
```

---

### 2.3 delete_access_form

```python
async def delete_access_form_tool(
    file_path: str,
    form_name: str,
    backup_first: bool = True
) -> str:
    """
    [PRO] Delete an Access form.

    Args:
        file_path: Path to .accdb file
        form_name: Name of form to delete
        backup_first: If True, export to temp before deleting

    Returns:
        Confirmation message
    """
    session = await get_or_create_session(file_path)
    app = session.app

    # Vérifier que le form existe
    form_exists = False
    for form in app.CurrentProject.AllForms:
        if form.Name.lower() == form_name.lower():
            form_exists = True
            break

    if not form_exists:
        raise ValueError(f"Form '{form_name}' not found")

    # Backup optionnel
    backup_path = None
    if backup_first:
        backup_dir = Path(file_path).parent / ".form_backups"
        backup_dir.mkdir(exist_ok=True)
        backup_path = backup_dir / f"{form_name}_{datetime.now():%Y%m%d_%H%M%S}.txt"
        app.SaveAsText(2, form_name, str(backup_path))  # acForm = 2

    # Supprimer
    app.DoCmd.DeleteObject(2, form_name)  # acForm = 2

    msg = f"## Form Deleted\n\n**Name:** {form_name}\n"
    if backup_path:
        msg += f"**Backup:** {backup_path}\n"

    return msg
```

---

### 2.4 export_form_definition

```python
async def export_form_definition_tool(
    file_path: str,
    form_name: str,
    output_path: Optional[str] = None
) -> str:
    """
    [PRO] Export Access form definition to text file (SaveAsText).

    This exports the complete form definition including:
    - All controls and their properties
    - Layout and positioning
    - VBA code behind the form
    - Event bindings

    The exported file can be:
    - Viewed to understand form structure
    - Modified with any text editor
    - Re-imported with import_form_definition
    - Version controlled with Git

    Args:
        file_path: Path to .accdb file
        form_name: Name of form to export
        output_path: Where to save .txt file (default: same folder as db)

    Returns:
        Path to exported file + preview of content
    """
    session = await get_or_create_session(file_path)
    app = session.app

    # Vérifier que le form existe
    form_exists = False
    for form in app.CurrentProject.AllForms:
        if form.Name.lower() == form_name.lower():
            form_exists = True
            actual_name = form.Name  # Nom exact avec casse
            break

    if not form_exists:
        raise ValueError(f"Form '{form_name}' not found")

    # Chemin de sortie
    if output_path:
        export_path = Path(output_path)
    else:
        export_path = Path(file_path).parent / f"{actual_name}.txt"

    # Export
    app.SaveAsText(2, actual_name, str(export_path))  # acForm = 2

    # Lire un aperçu
    with open(export_path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
        preview = content[:2000]
        if len(content) > 2000:
            preview += "\n... [truncated]"

    return f"""## Form Exported

**Form:** {actual_name}
**Output:** {export_path}
**Size:** {len(content)} characters

### Preview

```
{preview}
```

You can now:
1. Read the full file with the Read tool
2. Modify the content
3. Re-import with `import_form_definition`
"""
```

---

### 2.5 import_form_definition

```python
async def import_form_definition_tool(
    file_path: str,
    form_name: str,
    definition_path: str,
    overwrite: bool = False
) -> str:
    """
    [PRO] Import Access form from text definition (LoadFromText).

    This imports a form definition previously exported with SaveAsText
    or manually created/modified.

    Args:
        file_path: Path to .accdb file
        form_name: Name for the imported form
        definition_path: Path to .txt definition file
        overwrite: If True, delete existing form first

    Returns:
        Success message

    Workflow example:
        1. export_form_definition("db.accdb", "frm_Old")
        2. Read and modify the .txt file
        3. import_form_definition("db.accdb", "frm_New", "modified.txt")
    """
    session = await get_or_create_session(file_path)
    app = session.app

    # Vérifier que le fichier source existe
    def_path = Path(definition_path)
    if not def_path.exists():
        raise FileNotFoundError(f"Definition file not found: {definition_path}")

    # Vérifier si le form existe déjà
    for form in app.CurrentProject.AllForms:
        if form.Name.lower() == form_name.lower():
            if overwrite:
                app.DoCmd.DeleteObject(2, form.Name)
            else:
                raise ValueError(
                    f"Form '{form_name}' already exists. "
                    "Use overwrite=True to replace."
                )

    # Import
    app.LoadFromText(2, form_name, str(def_path))  # acForm = 2

    return f"""## Form Imported

**Form:** {form_name}
**Source:** {def_path}
**Overwrite:** {overwrite}

Form imported successfully. You can now:
- Open it in Access to verify
- Use `list_access_forms` to confirm
- Use `export_form_definition` to re-export
"""
```

---

## Phase 3: Intégration Server

### 3.1 Ajouter au server.py

**Fichier:** `packages/pro/src/vba_mcp_pro/server.py`

```python
# Imports (ajouter)
from .tools import (
    # ... existing ...
    list_access_forms_tool,
    create_access_form_tool,
    delete_access_form_tool,
    export_form_definition_tool,
    import_form_definition_tool,
)

# Tool definitions (ajouter dans TOOLS list)
{
    "name": "list_access_forms",
    "description": "[PRO] List all forms in an Access database.",
    "inputSchema": {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "Path to .accdb file"
            }
        },
        "required": ["file_path"]
    }
},
{
    "name": "create_access_form",
    "description": "[PRO] Create a new Access form. Can auto-generate controls from table fields.",
    "inputSchema": {
        "type": "object",
        "properties": {
            "file_path": {"type": "string"},
            "form_name": {"type": "string"},
            "record_source": {"type": "string", "description": "Table or query to bind to"},
            "auto_generate": {"type": "boolean", "default": False},
            "form_type": {"type": "string", "enum": ["single", "continuous", "datasheet"]}
        },
        "required": ["file_path", "form_name"]
    }
},
{
    "name": "delete_access_form",
    "description": "[PRO] Delete an Access form.",
    "inputSchema": {
        "type": "object",
        "properties": {
            "file_path": {"type": "string"},
            "form_name": {"type": "string"},
            "backup_first": {"type": "boolean", "default": True}
        },
        "required": ["file_path", "form_name"]
    }
},
{
    "name": "export_form_definition",
    "description": "[PRO] Export Access form to text file. Enables viewing/editing form structure as text.",
    "inputSchema": {
        "type": "object",
        "properties": {
            "file_path": {"type": "string"},
            "form_name": {"type": "string"},
            "output_path": {"type": "string", "description": "Optional custom output path"}
        },
        "required": ["file_path", "form_name"]
    }
},
{
    "name": "import_form_definition",
    "description": "[PRO] Import Access form from text definition file.",
    "inputSchema": {
        "type": "object",
        "properties": {
            "file_path": {"type": "string"},
            "form_name": {"type": "string"},
            "definition_path": {"type": "string"},
            "overwrite": {"type": "boolean", "default": False}
        },
        "required": ["file_path", "form_name", "definition_path"]
    }
}

# Handler (ajouter dans handle_call_tool)
elif name == "list_access_forms":
    result = await list_access_forms_tool(arguments["file_path"])
elif name == "create_access_form":
    result = await create_access_form_tool(
        arguments["file_path"],
        arguments["form_name"],
        arguments.get("record_source"),
        arguments.get("auto_generate", False),
        arguments.get("form_type", "single")
    )
elif name == "delete_access_form":
    result = await delete_access_form_tool(
        arguments["file_path"],
        arguments["form_name"],
        arguments.get("backup_first", True)
    )
elif name == "export_form_definition":
    result = await export_form_definition_tool(
        arguments["file_path"],
        arguments["form_name"],
        arguments.get("output_path")
    )
elif name == "import_form_definition":
    result = await import_form_definition_tool(
        arguments["file_path"],
        arguments["form_name"],
        arguments["definition_path"],
        arguments.get("overwrite", False)
    )
```

---

### 3.2 Mettre à jour __init__.py

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/__init__.py`

```python
from .office_automation import (
    # ... existing ...
    list_access_forms_tool,
    create_access_form_tool,
    delete_access_form_tool,
    export_form_definition_tool,
    import_form_definition_tool,
)

__all__ = [
    # ... existing ...
    "list_access_forms_tool",
    "create_access_form_tool",
    "delete_access_form_tool",
    "export_form_definition_tool",
    "import_form_definition_tool",
]
```

---

## Phase 4: Tests

### 4.1 Tests unitaires

**Fichier:** `tests/test_access_forms.py`

```python
import pytest
from pathlib import Path

ACCESS_FILE = Path(__file__).parent.parent / "vba-mcp-demo/sample-files/demo-database.accdb"

class TestListForms:
    async def test_list_forms_empty_db(self):
        """Test listing forms in db with no forms."""
        result = await list_access_forms_tool(str(ACCESS_FILE))
        assert "Access Forms" in result

    async def test_list_forms_with_forms(self):
        """Test listing forms after creating one."""
        # Create form first
        await create_access_form_tool(str(ACCESS_FILE), "frm_Test")
        result = await list_access_forms_tool(str(ACCESS_FILE))
        assert "frm_Test" in result
        # Cleanup
        await delete_access_form_tool(str(ACCESS_FILE), "frm_Test")


class TestCreateForm:
    async def test_create_empty_form(self):
        result = await create_access_form_tool(
            str(ACCESS_FILE),
            "frm_Empty"
        )
        assert "Form Created" in result
        # Cleanup
        await delete_access_form_tool(str(ACCESS_FILE), "frm_Empty")

    async def test_create_bound_form(self):
        result = await create_access_form_tool(
            str(ACCESS_FILE),
            "frm_Employees",
            record_source="Employees"
        )
        assert "Employees" in result
        await delete_access_form_tool(str(ACCESS_FILE), "frm_Employees")

    async def test_create_auto_generated(self):
        result = await create_access_form_tool(
            str(ACCESS_FILE),
            "frm_AutoGen",
            record_source="Employees",
            auto_generate=True
        )
        assert "Auto-generated: Yes" in result
        await delete_access_form_tool(str(ACCESS_FILE), "frm_AutoGen")

    async def test_create_duplicate_fails(self):
        await create_access_form_tool(str(ACCESS_FILE), "frm_Dup")
        with pytest.raises(ValueError, match="already exists"):
            await create_access_form_tool(str(ACCESS_FILE), "frm_Dup")
        await delete_access_form_tool(str(ACCESS_FILE), "frm_Dup")


class TestDeleteForm:
    async def test_delete_form(self):
        await create_access_form_tool(str(ACCESS_FILE), "frm_ToDelete")
        result = await delete_access_form_tool(str(ACCESS_FILE), "frm_ToDelete")
        assert "Form Deleted" in result

    async def test_delete_with_backup(self):
        await create_access_form_tool(str(ACCESS_FILE), "frm_Backup")
        result = await delete_access_form_tool(
            str(ACCESS_FILE),
            "frm_Backup",
            backup_first=True
        )
        assert "Backup:" in result

    async def test_delete_nonexistent_fails(self):
        with pytest.raises(ValueError, match="not found"):
            await delete_access_form_tool(str(ACCESS_FILE), "frm_DoesNotExist")


class TestExportForm:
    async def test_export_form(self):
        # Create form to export
        await create_access_form_tool(str(ACCESS_FILE), "frm_Export")

        result = await export_form_definition_tool(
            str(ACCESS_FILE),
            "frm_Export"
        )

        assert "Form Exported" in result
        assert "Preview" in result

        # Cleanup
        await delete_access_form_tool(str(ACCESS_FILE), "frm_Export")

    async def test_export_custom_path(self, tmp_path):
        await create_access_form_tool(str(ACCESS_FILE), "frm_CustomPath")

        output = tmp_path / "custom_export.txt"
        result = await export_form_definition_tool(
            str(ACCESS_FILE),
            "frm_CustomPath",
            str(output)
        )

        assert output.exists()
        await delete_access_form_tool(str(ACCESS_FILE), "frm_CustomPath")


class TestImportForm:
    async def test_import_form(self, tmp_path):
        # Create and export
        await create_access_form_tool(str(ACCESS_FILE), "frm_Original")
        export_path = tmp_path / "form.txt"
        await export_form_definition_tool(
            str(ACCESS_FILE),
            "frm_Original",
            str(export_path)
        )

        # Import as new form
        result = await import_form_definition_tool(
            str(ACCESS_FILE),
            "frm_Imported",
            str(export_path)
        )

        assert "Form Imported" in result

        # Cleanup
        await delete_access_form_tool(str(ACCESS_FILE), "frm_Original")
        await delete_access_form_tool(str(ACCESS_FILE), "frm_Imported")

    async def test_import_overwrite(self, tmp_path):
        # Create and export
        await create_access_form_tool(str(ACCESS_FILE), "frm_Overwrite")
        export_path = tmp_path / "form.txt"
        await export_form_definition_tool(
            str(ACCESS_FILE),
            "frm_Overwrite",
            str(export_path)
        )

        # Import with overwrite
        result = await import_form_definition_tool(
            str(ACCESS_FILE),
            "frm_Overwrite",
            str(export_path),
            overwrite=True
        )

        assert "Form Imported" in result

        # Cleanup
        await delete_access_form_tool(str(ACCESS_FILE), "frm_Overwrite")


class TestWorkflow:
    """Test complete workflow: create -> export -> modify -> import."""

    async def test_full_workflow(self, tmp_path):
        # 1. Create form
        await create_access_form_tool(
            str(ACCESS_FILE),
            "frm_Workflow",
            record_source="Employees",
            auto_generate=True
        )

        # 2. Export
        export_path = tmp_path / "workflow.txt"
        await export_form_definition_tool(
            str(ACCESS_FILE),
            "frm_Workflow",
            str(export_path)
        )

        # 3. Read and modify (simulate Claude editing)
        with open(export_path, 'r') as f:
            content = f.read()

        # Add a button (simplified)
        modified = content.replace(
            "End\nEnd",
            """    Begin CommandButton
            Left =500
            Top =2000
            Width =2000
            Height =400
            Name ="btnSave"
            Caption ="Save"
        End
    End
End"""
        )

        modified_path = tmp_path / "workflow_modified.txt"
        with open(modified_path, 'w') as f:
            f.write(modified)

        # 4. Import modified version
        result = await import_form_definition_tool(
            str(ACCESS_FILE),
            "frm_Workflow_V2",
            str(modified_path)
        )

        assert "Form Imported" in result

        # 5. Cleanup
        await delete_access_form_tool(str(ACCESS_FILE), "frm_Workflow")
        await delete_access_form_tool(str(ACCESS_FILE), "frm_Workflow_V2")
```

---

## Phase 5: Documentation

### 5.1 Mettre à jour CLAUDE.md

Ajouter dans la section "Outils MCP Disponibles":

```markdown
#### Access Forms (NEW - v0.7.0)
- `list_access_forms` - Lister formulaires
- `create_access_form` - Créer formulaire (vide ou auto-généré)
- `delete_access_form` - Supprimer formulaire
- `export_form_definition` - Exporter en texte (SaveAsText)
- `import_form_definition` - Importer depuis texte (LoadFromText)
```

### 5.2 Créer guide d'utilisation

**Fichier:** `docs/ACCESS_FORMS_GUIDE.md`

```markdown
# Guide: Formulaires Access

## Workflow recommandé

### Créer un formulaire CRUD automatique

1. Utiliser `create_access_form` avec `auto_generate=True`
2. Exporter avec `export_form_definition`
3. Modifier le texte pour ajustements
4. Réimporter avec `import_form_definition`

### Modifier un formulaire existant

1. Exporter avec `export_form_definition`
2. Lire le fichier avec l'outil Read
3. Modifier le contenu
4. Réimporter avec `overwrite=True`

## Format SaveAsText

Le format texte Access a cette structure:
- Version header
- Begin Form ... End block
- Control blocks (Label, TextBox, CommandButton, etc.)
- CodeBehindForm section (VBA code)
```

---

## Critères de Succès

### Phase 1 Complete ✅
- [x] Format SaveAsText documenté (UTF-16, structure Begin/End)
- [x] Tests manuels SaveAsText/LoadFromText réussis

### Phase 2 Complete ✅
- [x] 5 outils implémentés
- [x] Tests manuels passent

### Phase 3 Complete ✅
- [x] Intégration server.py
- [x] Exports dans __init__.py

### Phase 4 Complete ✅
- [x] Tests d'intégration passent (create → export → import → delete)

### Phase 5 Complete ✅
- [x] CLAUDE.md mis à jour
- [x] TODO.md mis à jour

---

## Notes Techniques

### Constantes Access

```python
# Object types
acForm = 2
acReport = 3
acMacro = 4
acModule = 5

# Control types
acLabel = 100
acRectangle = 101
acLine = 102
acImage = 103
acCommandButton = 104
acOptionButton = 105
acCheckBox = 106
acOptionGroup = 107
acBoundObjectFrame = 108
acTextBox = 109
acListBox = 110
acComboBox = 111
acSubform = 112
acObjectFrame = 114
acPageBreak = 118
acPage = 124
acCustomControl = 119
acToggleButton = 122
acTabCtl = 123
acWebBrowser = 128
acNavigationControl = 129
acNavigationButton = 130
acAttachment = 126
```

### Sections de formulaire

```python
# Form sections
acDetail = 0
acHeader = 1
acFooter = 2
acPageHeader = 3
acPageFooter = 4
acGroupLevel1Header = 5
acGroupLevel1Footer = 6
# etc.
```

---

**Auteur:** Claude Code
**Date:** 2025-12-30
**Status:** En cours
