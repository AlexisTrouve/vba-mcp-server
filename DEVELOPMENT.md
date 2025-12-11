# Development Guide

Guide complet pour développer et tester le monorepo VBA MCP.

## Structure du Monorepo

```
vba-mcp-monorepo/
├── packages/
│   ├── core/          # Bibliothèque partagée (extraction VBA)
│   ├── lite/          # Serveur MCP gratuit (lecture seule)
│   └── pro/           # Serveur MCP commercial (lecture + écriture)
├── tests/             # Tests communs
├── docs/              # Documentation
└── examples/          # Exemples d'utilisation
```

## Prérequis

- **Python 3.8+** (recommandé : 3.10+)
- **Git**
- **Windows uniquement** pour tester l'injection VBA (package Pro)
- **Microsoft Office** installé pour les tests réels

## Installation - Développement Local

### 1. Cloner le repository

```bash
git clone git@git.etheryale.com:StillHammer/vba-mcp-server-pro.git
cd vba-mcp-monorepo
```

### 2. Créer un environnement virtuel

```bash
# Linux/macOS
python3 -m venv venv
source venv/bin/activate

# Windows
python -m venv venv
venv\Scripts\activate
```

### 3. Installer les packages en mode editable

**Important** : Installer dans cet ordre (dépendances) :

```bash
# 1. Core (pas de dépendances internes)
pip install -e ./packages/core

# 2. Lite (dépend de core)
pip install -e ./packages/lite

# 3. Pro (dépend de core)
# Sur Windows uniquement :
pip install -e "./packages/pro[windows]"

# Sur Linux/macOS (sans pywin32) :
pip install -e ./packages/pro
```

### 4. Vérifier l'installation

```bash
# Tester les imports
python -c "from vba_mcp_core import OfficeHandler; print('✓ Core OK')"
python -c "from vba_mcp_lite.server import run; print('✓ Lite OK')"
python -c "from vba_mcp_pro.server import run; print('✓ Pro OK')"

# Vérifier les commandes
vba-mcp-server --help
vba-mcp-server-pro --help
```

## Structure des Packages

### Core (`vba-mcp-core`)

**Bibliothèque partagée** - MIT License

```
packages/core/
├── src/vba_mcp_core/
│   ├── __init__.py
│   ├── lib/
│   │   ├── office_handler.py    # Extraction VBA
│   │   └── vba_parser.py        # Analyse syntaxique
│   └── tools/
│       ├── extract.py           # Tool: extraire code
│       ├── list_modules.py      # Tool: lister modules
│       └── analyze.py           # Tool: analyser structure
└── pyproject.toml
```

**Dépendances** :
- `oletools>=0.60` - Extraction VBA
- `openpyxl>=3.0.0` - Manipulation Excel

### Lite (`vba-mcp-server`)

**Serveur MCP gratuit** - MIT License

```
packages/lite/
├── src/vba_mcp_lite/
│   ├── __init__.py
│   └── server.py               # Serveur MCP
└── pyproject.toml
```

**Features** :
- ✅ Extraire code VBA
- ✅ Lister modules
- ✅ Analyser structure

**Dépendances** :
- `vba-mcp-core>=0.1.0`
- `mcp>=1.0.0`

### Pro (`vba-mcp-server-pro`)

**Serveur MCP commercial** - Proprietary License

```
packages/pro/
├── src/vba_mcp_pro/
│   ├── __init__.py
│   ├── server.py               # Serveur MCP Pro
│   └── tools/
│       ├── inject.py           # Injection VBA (Windows COM)
│       ├── refactor.py         # Suggestions de refactoring
│       └── backup.py           # Gestion des backups
└── pyproject.toml
```

**Features** :
- ✅ Toutes les features Lite
- ✅ **Injection VBA** (Windows + pywin32)
- ✅ **Refactoring AI**
- ✅ **Backup management**

**Dépendances** :
- `vba-mcp-core>=0.1.0`
- `mcp>=1.0.0`
- `pywin32>=300` (Windows uniquement)

## Développement

### Workflow

1. **Faire vos modifications** dans le package concerné
2. **Tester localement** (imports, syntax)
3. **Commit & push**
4. **Publier sur PyPI** (voir section Publication)

### Tests rapides

```bash
# Vérifier la syntaxe de tous les fichiers Python
find packages/ -name "*.py" -exec python -m py_compile {} \;

# Tester les imports
PYTHONPATH=packages/core/src:packages/pro/src python -c "
from vba_mcp_core import OfficeHandler
from vba_mcp_pro.tools import inject_vba_tool
print('✓ All imports OK')
"
```

### Tests unitaires

```bash
# Installer pytest
pip install pytest pytest-asyncio

# Lancer les tests
pytest tests/
```

## Configuration MCP

### Claude Desktop - Lite Version

Ajouter dans `claude_desktop_config.json` :

```json
{
  "mcpServers": {
    "vba": {
      "command": "vba-mcp-server"
    }
  }
}
```

### Claude Desktop - Pro Version

```json
{
  "mcpServers": {
    "vba-pro": {
      "command": "vba-mcp-server-pro"
    }
  }
}
```

## Tester l'Injection VBA (Pro - Windows)

### Prérequis

1. **Windows** uniquement
2. **Microsoft Office** installé (Excel, Word, ou Access)
3. **Activer "Trust access to VBA project object model"** :
   - Excel : File > Options > Trust Center > Trust Center Settings > Macro Settings
   - Cocher "Trust access to the VBA project object model"

### Test manuel

```python
import asyncio
from pathlib import Path
from vba_mcp_pro.tools import inject_vba_tool

async def test_injection():
    code = """
Sub HelloWorld()
    MsgBox "Hello from Python!"
End Sub
"""
    result = await inject_vba_tool(
        file_path="C:\\path\\to\\test.xlsm",
        module_name="TestModule",
        code=code,
        create_backup=True
    )
    print(result)

asyncio.run(test_injection())
```

### Vérifier le backup

Les backups sont créés dans `.vba_backups/` :

```
C:\path\to\
├── test.xlsm
└── .vba_backups/
    ├── test_backup_20241211_153045.xlsm
    └── .vba_backups.json
```

## Publication sur PyPI

### Prérequis

```bash
pip install build twine
```

### 1. Publier Core (en premier)

```bash
cd packages/core
python -m build
twine upload dist/*
```

### 2. Publier Lite

```bash
cd packages/lite
python -m build
twine upload dist/*
```

### 3. Publier Pro

```bash
cd packages/pro
python -m build
twine upload dist/*
```

### Vérifier la publication

```bash
# Installer depuis PyPI
pip install vba-mcp-server-pro[windows]

# Tester
vba-mcp-server-pro --help
```

## Troubleshooting

### Import Error: No module named 'mcp'

```bash
pip install mcp
```

### Import Error: No module named 'oletools'

```bash
pip install oletools
```

### Windows: pywin32 not found

```bash
pip install pywin32
```

### WSL: Cannot create venv

```bash
sudo apt install python3-venv python3-full
```

### VBA Injection: Access Denied

Vérifier dans Excel/Word :
- File > Options > Trust Center > Trust Center Settings > Macro Settings
- Cocher "Trust access to the VBA project object model"

## Commandes Utiles

```bash
# Nettoyer les fichiers de build
find . -type d -name "__pycache__" -exec rm -rf {} +
find . -type d -name "*.egg-info" -exec rm -rf {} +
find . -type d -name "dist" -exec rm -rf {} +
find . -type d -name "build" -exec rm -rf {} +

# Mettre à jour les versions
# Éditer pyproject.toml dans chaque package

# Rebuild après modifications
cd packages/core && python -m build
cd packages/lite && python -m build
cd packages/pro && python -m build
```

## Architecture MCP

### Flow d'utilisation

```
User (Claude Desktop)
    |
    | MCP Protocol (stdio)
    v
MCP Server (vba-mcp-server-pro)
    |
    | Tool Calls
    v
VBA Tools (inject, refactor, backup)
    |
    | Core Library
    v
Office Files (.xlsm, .docm, .accdb)
```

### Cycle de vie d'une requête

1. **User** : "Modifie la fonction Calculate dans budget.xlsm"
2. **Claude** : Analyse et extrait d'abord le code avec `extract_vba`
3. **Claude** : Génère le nouveau code VBA
4. **Claude** : Appelle `inject_vba` avec le nouveau code
5. **Server** : Crée un backup automatique
6. **Server** : Ouvre Excel via COM (pywin32)
7. **Server** : Injecte le code dans le module
8. **Server** : Sauvegarde et ferme
9. **Claude** : Confirme à l'utilisateur

## Support

- **Email** : alexistrouve.pro@gmail.com
- **Git** : git@git.etheryale.com:StillHammer/vba-mcp-server-pro.git

## Licence

- **Core** : MIT License
- **Lite** : MIT License
- **Pro** : Commercial License (contact pour licensing)
