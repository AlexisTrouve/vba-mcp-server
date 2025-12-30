# VBA MCP Server - Instructions pour Claude

## Projet
VBA MCP Server - Serveur Model Context Protocol pour manipulation de code VBA dans les fichiers Microsoft Office.

## Statut Actuel : v0.6.0 - Production Ready

**Dernière mise à jour:** 30 décembre 2024

### Résultats des Tests

| Composant | Tests | Status |
|-----------|-------|--------|
| Excel | 20/20 (100%) | Stable |
| Access | 13/13 (100%) | Stable |
| VBA Injection | 100% | Stable |
| VBA Validation | 100% | Stable |

### Fonctionnalités Opérationnelles

| Fonctionnalité | Excel | Access | Notes |
|----------------|-------|--------|-------|
| Extraction VBA | Yes | Yes (COM) | |
| Injection VBA | Yes | Yes | |
| Liste modules | Yes | Yes (COM) | oletools ne supporte pas .accdb |
| Exécution macros | Yes | Limited | Access macros != VBA |
| Validation syntaxe | Yes | Yes | |
| Lecture données | Yes | Yes | |
| Écriture données | Yes | Yes | append/replace |
| SQL SELECT | - | Yes | |
| SQL INSERT/UPDATE/DELETE | - | Yes | NEW |
| Excel Tables | Yes | - | 6 outils |
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

## Outils MCP Disponibles (24 total)

### Lite (6 outils)
- `extract_vba` - Extraire code VBA
- `list_modules` - Lister modules VBA
- `analyze_structure` - Analyser structure du code

### Pro (18 outils)

#### VBA
- `inject_vba` - Injecter code VBA
- `validate_vba` - Valider syntaxe VBA
- `run_macro` - Exécuter macros

#### Données
- `open_in_office` - Ouvrir fichier Office
- `get_worksheet_data` - Lire données (Excel/Access)
- `set_worksheet_data` - Écrire données (Excel/Access)

#### Excel Tables
- `list_excel_tables` - Lister tables Excel
- `create_excel_table` - Créer table Excel
- `insert_rows` / `delete_rows` - Gérer lignes
- `insert_columns` / `delete_columns` - Gérer colonnes

#### Access
- `list_access_tables` - Lister tables avec schéma
- `list_access_queries` - Lister requêtes sauvegardées
- `run_access_query` - Exécuter SQL (SELECT/INSERT/UPDATE/DELETE)

#### Backup
- `create_backup` - Créer sauvegarde
- `restore_backup` - Restaurer sauvegarde
- `list_backups` - Lister sauvegardes

## Fichiers Clés

| Fichier | Description |
|---------|-------------|
| `packages/pro/src/vba_mcp_pro/server.py` | Serveur MCP principal |
| `packages/pro/src/vba_mcp_pro/tools/office_automation.py` | Outils Access + Excel |
| `packages/pro/src/vba_mcp_pro/tools/inject.py` | Injection VBA |
| `packages/pro/src/vba_mcp_pro/tools/validate.py` | Validation syntaxe |
| `packages/pro/src/vba_mcp_pro/session_manager.py` | Gestion sessions Office |
| `test_access_complete.py` | Suite de tests Access |
| `TODO.md` | Roadmap fonctionnalités |

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
- Forms/Reports Access

## Contact

Alexis Trouve - alexistrouve.pro@gmail.com
