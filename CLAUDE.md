# VBA MCP Server - Instructions pour Claude

## Projet
VBA MCP Server - Serveur Model Context Protocol pour manipulation de code VBA dans les fichiers Microsoft Office.

## Statut Actuel : v0.8.0 - Production Ready

**Derniere mise a jour:** 30 decembre 2024

### Resultats des Tests

| Composant | Tests | Status |
|-----------|-------|--------|
| Excel | 20/20 (100%) | Stable |
| Access | 16/16 (100%) | Stable |
| VBA Injection | 100% | Stable |
| VBA Validation | 100% | Stable |
| Access VBA COM | 3/3 (100%) | NEW |

### Fonctionnalites Operationnelles

| Fonctionnalite | Excel | Access | Notes |
|----------------|-------|--------|-------|
| Extraction VBA | Yes | Yes (COM) | extract_vba_access pour .accdb |
| Injection VBA | Yes | Yes | |
| Liste modules | Yes | Yes (COM) | oletools ne supporte pas .accdb |
| Analyse structure | Yes | Yes (COM) | analyze_structure_access NEW |
| Compilation VBA | Yes | Yes | compile_vba NEW |
| Execution macros | Yes | Limited | Access macros != VBA |
| Validation syntaxe | Yes | Yes | |
| Lecture donnees | Yes | Yes | |
| Ecriture donnees | Yes | Yes | append/replace |
| SQL SELECT | - | Yes | |
| SQL INSERT/UPDATE/DELETE | - | Yes | |
| Excel Tables | Yes | - | 6 outils |
| Access Forms | - | Yes | 5 outils |
| Backup/Rollback | Yes | Yes | |

## Structure du Projet

```
vba-mcp-monorepo/
├── packages/
│   ├── core/       # Bibliothèque partagée (MIT)
│   ├── lite/       # Serveur open source (MIT)
│   └── pro/        # Version commerciale
├── docs/           # Documentation
├── tests/          # Tests d'intégration
├── plans/          # Plans d'implémentation
├── TODO.md         # Roadmap des fonctionnalités futures
└── CLAUDE.md       # Ce fichier
```

## Commandes Utiles

### Lancer le serveur MCP
```bash
# Avec PYTHONPATH
set PYTHONPATH=packages/core/src;packages/lite/src;packages/pro/src
python -m vba_mcp_pro.server

# Ou via script
start_vba_mcp.bat
```

### Lancer les tests
```bash
# Tests unitaires
pytest packages/pro/tests/

# Test suite Access complète
python test_access_complete.py

# Test suite Access basique
python test_access_tools.py
```

## Outils MCP Disponibles (32 total)

### Lite (3 outils)
- `extract_vba` - Extraire code VBA (oletools)
- `list_modules` - Lister modules VBA
- `analyze_structure` - Analyser structure du code

### Pro (29 outils)

#### VBA
- `inject_vba` - Injecter code VBA
- `validate_vba` - Valider syntaxe VBA
- `run_macro` - Executer macros

#### Donnees
- `open_in_office` - Ouvrir fichier Office
- `get_worksheet_data` - Lire donnees (Excel/Access)
- `set_worksheet_data` - Ecrire donnees (Excel/Access)

#### Excel Tables
- `list_tables` - Lister tables Excel
- `create_table` - Creer table Excel
- `insert_rows` / `delete_rows` - Gerer lignes
- `insert_columns` / `delete_columns` - Gerer colonnes

#### Access Data
- `list_access_tables` - Lister tables avec schema
- `list_access_queries` - Lister requetes sauvegardees
- `run_access_query` - Executer SQL (SELECT/INSERT/UPDATE/DELETE)

#### Access Forms
- `list_access_forms` - Lister tous les formulaires
- `create_access_form` - Creer formulaire (vide ou lie a table)
- `delete_access_form` - Supprimer formulaire (avec backup)
- `export_form_definition` - Exporter en texte (SaveAsText)
- `import_form_definition` - Importer depuis texte (LoadFromText)

#### Access VBA via COM (NEW in v0.8.0)
- `extract_vba_access` - Extraire VBA depuis .accdb via COM
- `analyze_structure_access` - Analyser structure VBA via COM
- `compile_vba` - Compiler projet VBA et detecter erreurs

#### Backup
- `create_backup` - Creer sauvegarde
- `restore_backup` - Restaurer sauvegarde
- `list_backups` - Lister sauvegardes

## Fichiers Cles

| Fichier | Description |
|---------|-------------|
| `packages/pro/src/vba_mcp_pro/server.py` | Serveur MCP principal |
| `packages/pro/src/vba_mcp_pro/tools/office_automation.py` | Outils Access + Excel |
| `packages/pro/src/vba_mcp_pro/tools/access_vba.py` | VBA Access via COM (NEW) |
| `packages/pro/src/vba_mcp_pro/tools/inject.py` | Injection VBA |
| `packages/pro/src/vba_mcp_pro/tools/validate.py` | Validation syntaxe |
| `packages/pro/src/vba_mcp_pro/session_manager.py` | Gestion sessions Office |
| `test_access_complete.py` | Suite de tests Access |
| `test_access_vba_tools.py` | Tests Access VBA COM (NEW) |
| `TODO.md` | Roadmap fonctionnalites |

## Limitations Connues

1. **Windows only** - Injection VBA nécessite pywin32 + COM
2. **Trust VBA** - Activer "Trust access to VBA project object model" dans Office
3. **oletools** - Ne supporte pas .accdb → Utiliser COM via session manager
4. **Access macros** - Les macros Access (UI) sont différentes des procédures VBA
5. **Sessions rapides** - 5+ injections en <2s peut crasher COM

## Projet Demo

Le projet `../vba-mcp-demo/sample-files/` contient :

### Excel
- `budget-analyzer.xlsm`
- `data-processor.xlsm`
- `report-generator.xlsm`
- `test-injection.xlsm`

### Access
- `demo-database.accdb` - Base de test avec :
  - 3 tables : Employees, Projects, ProjectAssignments
  - 4 queries : qryITEmployees, qryActiveProjects, etc.
  - Modules VBA : DemoModule, InjectedModule

## Roadmap (TODO.md)

### Haute Priorité
- Requêtes paramétrées (sécurité anti-injection SQL)
- Pivot Tables Excel
- Named Ranges

### Moyenne Priorité
- CREATE/ALTER/DROP TABLE pour Access
- Transactions (BEGIN/COMMIT/ROLLBACK)
- Charts Excel

### Basse Priorité
- Support Word/PowerPoint
- Reports Access (export PDF)

## Contact

Alexis Trouve - alexistrouve.pro@gmail.com
