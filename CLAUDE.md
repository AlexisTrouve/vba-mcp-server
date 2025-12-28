# VBA MCP Server - Instructions pour Claude

## Projet
VBA MCP Server - Serveur Model Context Protocol pour manipulation de code VBA dans les fichiers Microsoft Office.

## Statut Actuel : v0.5.0 - Production Ready

**Tous les tests passent (100%)** - Session 3 du 28 décembre 2025

### Fonctionnalités Opérationnelles

| Fonctionnalité | Status | Notes |
|----------------|--------|-------|
| Extraction VBA | 100% | Stable |
| Liste modules/macros | 100% | Stable |
| Exécution macros | 100% | Stable |
| **Injection VBA** | **100%** | Corrigé v0.5.0 |
| Validation syntaxe | 100% | Amélioré v0.5.0 |
| Excel Tables | 100% | Stable |
| Backup/Rollback | 100% | Stable |

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

## Corrections Récentes (v0.5.0)

### 1. Injection VBA - CORRIGÉ
- Suppression CoInitialize/CoUninitialize redondants
- Normalisation du code pour comparaison robuste
- Plus d'erreur "Code mismatch"

### 2. Validation Syntaxique - AMÉLIORÉ
- Détection blocs non fermés (If/For/While/Do/With/Select/Sub/Function)
- Gestion If single-ligne vs multi-ligne
- Messages d'erreur précis

## Fichiers Clés

- `packages/pro/src/vba_mcp_pro/tools/inject.py` - Injection VBA
- `packages/pro/src/vba_mcp_pro/tools/validate.py` - Validation
- `packages/pro/src/vba_mcp_pro/session_manager.py` - Sessions Office
- `TEST_RESULTS_SESSION3.md` - Résultats tests 100%

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
