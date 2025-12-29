# VBA MCP Server - Instructions pour Claude

## Projet
VBA MCP Server - Serveur Model Context Protocol pour manipulation de code VBA dans les fichiers Microsoft Office.

## Statut Actuel : v0.6.0 - Production Ready

**Tous les tests passent (100%)** - Session 3 du 28 décembre 2025

### Fonctionnalités Opérationnelles

| Fonctionnalité | Status | Notes |
|----------------|--------|-------|
| Extraction VBA | 100% | Excel + Access |
| Liste modules/macros | 100% | Excel + Access |
| Exécution macros | 100% | Excel + Access |
| **Injection VBA** | **100%** | Excel + Access |
| Validation syntaxe | 100% | Excel + Access |
| Excel Tables | 100% | Stable |
| Backup/Rollback | 100% | Stable |
| **Access Data** | **100%** | NEW v0.6.0 |
| **Access Queries** | **100%** | NEW v0.6.0 |

## Structure du Projet

```
vba-mcp-monorepo/
├── packages/
│   ├── core/       # Bibliothèque partagée (MIT)
│   ├── lite/       # Serveur open source (MIT)
│   └── pro/        # Version commerciale
├── docs/           # Documentation
└── tests/          # Tests d'intégration
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

### Tester l'injection VBA
```bash
python test_one_injection.py
```

### Lancer les tests
```bash
pytest packages/pro/tests/
```

## Nouveautés v0.6.0 - Support Access Complet

### Nouveaux Outils Access
- `list_access_tables` - Liste tables avec schéma (champs, types, nb records)
- `list_access_queries` - Liste requêtes sauvegardées (QueryDefs)
- `run_access_query` - Exécute requêtes ou SQL direct
- `get_worksheet_data` - Amélioré avec filtres SQL pour Access
- `set_worksheet_data` - Supporte écriture Access (append/replace)

### Fonctionnalités Access
- Lecture données avec WHERE, ORDER BY, LIMIT
- Écriture données en mode append ou replace
- Exécution de requêtes sauvegardées
- SQL personnalisé direct
- Schéma complet des tables

## Corrections v0.5.0

### 1. Injection VBA - CORRIGÉ
- Suppression CoInitialize/CoUninitialize redondants
- Normalisation du code pour comparaison robuste
- Plus d'erreur "Code mismatch"

### 2. Validation Syntaxique - AMÉLIORÉ
- Détection blocs non fermés (If/For/While/Do/With/Select/Sub/Function)
- Gestion If single-ligne vs multi-ligne
- Messages d'erreur précis

## Fichiers Clés

- `packages/pro/src/vba_mcp_pro/tools/office_automation.py` - Outils Access + Excel
- `packages/pro/src/vba_mcp_pro/tools/inject.py` - Injection VBA
- `packages/pro/src/vba_mcp_pro/tools/validate.py` - Validation
- `packages/pro/src/vba_mcp_pro/session_manager.py` - Sessions Office
- `plans/ACCESS_SUPPORT_PLAN.md` - Plan implémentation Access

## Limitations Connues

1. **Sessions rapides** : Enchaîner 5+ injections en <2s peut crasher (COM)
2. **Windows only** : Injection nécessite pywin32 + COM
3. **Trust VBA** : Excel doit avoir "Trust access to VBA project object model"

## Projet Demo

Le projet `../vba-mcp-demo/` contient des fichiers Excel de test :
- `budget-analyzer.xlsm`
- `data-processor.xlsm`
- `report-generator.xlsm`

## Contact

Alexis Trouve - alexistrouve.pro@gmail.com
