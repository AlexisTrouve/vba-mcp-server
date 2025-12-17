# ✅ Checklist de Test - VBA MCP Pro v0.3.0

**Date:** 2025-12-14
**Version:** 0.3.0 (21 outils, +6 nouveaux pour Excel Tables)

---

## 📋 Avant de Tester

### 1. Vérifier l'Installation

- [ ] Python installé (Windows)
- [ ] Microsoft Office installé (Excel, Word ou Access)
- [ ] pywin32 installé : `pip install pywin32`
- [ ] MCP SDK installé : `pip install mcp`

### 2. Configuration Claude Desktop

**Fichier:** `%APPDATA%\Claude\claude_desktop_config.json`

Utilise cette configuration :

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

**⚠️ Attention:** Adapte les chemins si nécessaire !

---

## 🔧 Test du Serveur (Windows uniquement)

### Option A: Script Batch (Recommandé)

1. Ouvre une invite de commande Windows
2. Navigue vers le dossier : `cd C:\Users\alexi\Documents\projects\vba-mcp-monorepo`
3. Lance : `test_server_standalone.bat`
4. Vérifie que tu vois : `[SUCCESS] Server is working!`
5. Vérifie le nombre d'outils : **21 tools**

### Option B: Script Python

1. Ouvre une invite de commande Windows
2. Navigue vers le dossier : `cd C:\Users\alexi\Documents\projects\vba-mcp-monorepo`
3. Lance : `python test_server.py`
4. Vérifie que tous les outils sont listés (21 total)

**Détail des outils attendus:**

- ✓ Core/Lite: 3 outils (extract_vba, list_modules, analyze_code)
- ✓ Pro - Injection/Validation: 3 outils (inject_vba, validate_vba_code, list_macros)
- ✓ Pro - Office Automation: 6 outils (open_in_office, run_macro, etc.)
- ✓ **Pro - Excel Tables (NEW!): 6 outils** (list_tables, insert_rows, delete_rows, insert_columns, delete_columns, create_table)
- ✓ Pro - Backup/Refactor: 2 outils (backup, refactor)

**TOTAL: 21 outils**

---

## 🚀 Redémarrer Claude Desktop

### Sur Windows :

1. **Ferme** complètement Claude Desktop (vérifier la barre des tâches)
2. **Relance** Claude Desktop
3. Attends 5-10 secondes que les serveurs MCP se connectent
4. Vérifie dans une nouvelle conversation :

```
What VBA MCP tools do you have available?
```

Tu devrais voir **21 tools** listés, incluant les nouveaux :
- list_tables
- insert_rows
- delete_rows
- insert_columns
- delete_columns
- create_table

---

## 🧪 Tests Rapides

### Test 1: Vérifier les Outils

```
List all VBA MCP tools you have, grouped by category
```

**Attendu:** 21 outils groupés en 5 catégories

---

### Test 2: Test Excel Tables (NOUVEAU!)

**Prérequis:** Ouvre `test.xlsm` dans Excel d'abord pour créer des données

```
In C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm:
1. Create an Excel table named "TestData" from range A1:C10 on Sheet1
2. List all tables to confirm it was created
3. Insert a new column "Status" at position D
4. Insert 3 rows at the end of the table
5. Get the data from the table to verify
```

**Attendu:** Toutes les opérations réussissent, le tableau est créé et modifié

---

### Test 3: VBA Validation (v0.2.0)

```
Validate this VBA code:
Sub Test()
    MsgBox "Hello World"
End Sub
```

**Attendu:** Validation réussie

---

### Test 4: Liste des Macros (v0.2.0)

```
List all macros in C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm
```

**Attendu:** Liste des macros publiques avec signatures

---

### Test 5: Workflow Complet (Office Automation)

```
1. Open C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm
2. List all open files
3. Get data from Sheet1, range A1:C5
4. Close the file and save
```

**Attendu:** Fichier s'ouvre dans Excel (visible), données lues, fichier fermé

---

## 🐛 Dépannage

### Serveur ne se connecte pas

1. Vérifie les logs Claude Desktop : `%APPDATA%\Claude\logs\mcp*.log`
2. Vérifie que Python est dans le PATH
3. Vérifie que pywin32 et mcp sont installés
4. Vérifie les chemins dans claude_desktop_config.json

### "Cannot run macro"

- Active "Trust access to the VBA project object model" dans Excel :
  File → Options → Trust Center → Trust Center Settings → Macro Settings

### Module not found

- Vérifie PYTHONPATH dans la config
- Vérifie que les 3 packages (core, lite, pro) existent

### Excel crash ou freeze

- Ferme Excel complètement avant de tester
- Désactive les compléments Excel
- Teste avec un fichier simple d'abord

---

## 📊 Résultats Attendus

### Serveur

- [x] Serveur s'importe sans erreur
- [x] 21 outils enregistrés
- [x] Syntaxe Python valide (vérifié)
- [x] Documentation complète

### Tests Fonctionnels (à faire sur Windows)

- [ ] Claude Desktop se connecte au serveur
- [ ] 21 outils visibles dans Claude
- [ ] Nouveaux outils Excel Tables fonctionnent
- [ ] Outils existants fonctionnent toujours
- [ ] Validation VBA fonctionne
- [ ] Office Automation fonctionne

---

## 📁 Fichiers de Test Disponibles

### Monorepo

- `test.xlsm` - Fichier de test principal

### Demo (vba-mcp-demo/sample-files/)

- `budget-analyzer.xlsm` - Analyse de budget
- `data-processor.xlsm` - Traitement de données
- `report-generator.xlsm` - Génération de rapports

---

## 📚 Documentation de Référence

- `QUICK_TEST_PROMPTS.md` - Prompts de test détaillés (avec nouveaux tests Excel Tables)
- `vba-mcp-demo/PROMPTS_READY_TO_USE.md` - Prompts prêts à l'emploi en français
- `packages/pro/CHANGELOG.md` - Historique des versions
- `packages/pro/README.md` - Documentation du package Pro

---

## ✅ Succès Final

Quand tout fonctionne, tu devrais pouvoir :

1. ✅ Voir 21 outils dans Claude Desktop
2. ✅ Créer des tableaux Excel avec `create_table`
3. ✅ Manipuler les lignes/colonnes des tableaux
4. ✅ Lire/écrire des données de tableaux structurés
5. ✅ Valider du code VBA avant injection
6. ✅ Exécuter des macros et voir Excel s'ouvrir
7. ✅ Faire tout ce qui était possible en v0.2.0 + les nouveautés v0.3.0

---

**Version du serveur:** 0.3.0
**Date:** 2025-12-14
**Nouveautés:** Support complet Excel Tables (6 nouveaux outils)
