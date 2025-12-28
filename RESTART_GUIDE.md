# 🔄 Guide de Redémarrage - VBA MCP Pro v0.3.0

## Pourquoi Redémarrer ?

Après avoir mis à jour le code du serveur VBA MCP Pro, tu dois redémarrer Claude Code pour que les changements soient pris en compte.

**Changements v0.3.0:**
- ✅ 6 nouveaux outils Excel Tables ajoutés
- ✅ Total: 21 outils (15 Pro + 6 Lite)

---

## 🚀 Étapes de Redémarrage (Windows)

### 1. Fermer Claude Code COMPLÈTEMENT

**Important:** Ne pas juste minimiser !

1. **Clic droit** sur l'icône Claude Code dans la barre des tâches
2. **Sélectionne** "Quitter" ou "Exit"
3. **Vérifier** dans la zone de notification (system tray) que Claude n'est pas en arrière-plan
4. **Si encore présent:** Ouvrir le Gestionnaire des tâches (Ctrl+Shift+Esc) et terminer le processus "Claude"

### 2. Attendre 5 Secondes

Laisse le temps au système de libérer les ressources.

### 3. Redémarrer Claude Code

1. **Clique** sur l'icône Claude Code pour relancer
2. **Attends** 5-10 secondes que le serveur MCP se connecte
3. **Cherche** l'icône marteau 🔨 en bas à droite (indique que le serveur MCP est connecté)

---

## ✅ Vérifier que Tout Fonctionne

### Test 1: Compter les Outils

Ouvre une nouvelle conversation dans Claude Code et tape:

```
What VBA MCP tools do you have available?
```

**Résultat attendu:**
- Claude doit lister **21 outils**
- Tu dois voir les nouveaux outils Excel Tables :
  - list_tables
  - insert_rows
  - delete_rows
  - insert_columns
  - delete_columns
  - create_table

### Test 2: Tester un Nouvel Outil

```
List all Excel tables in C:\Users\alexi\Documents\projects\vba-mcp-monorepo\test.xlsm
```

**Résultat attendu:**
- Si aucun tableau n'existe encore : "No tables found"
- Si des tableaux existent : Liste avec noms, dimensions, en-têtes

---

## 🐛 Dépannage

### Le serveur ne se connecte pas (pas d'icône 🔨)

**Vérifier les logs:**
1. Dans Claude Code: **%USERPROFILE%\.claude\logs\mcp*.log**
2. Cherche les erreurs dans les logs MCP
3. Vérifie particulièrement les lignes avec `vba-mcp-pro`

**Solutions:**

**Erreur "Module not found"**
```
Vérifie le PYTHONPATH dans:
%USERPROFILE%\.claude\config.json

Doit contenir:
"PYTHONPATH": "C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo\\packages\\core\\src;C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo\\packages\\lite\\src;C:\\Users\\alexi\\Documents\\projects\\vba-mcp-monorepo\\packages\\pro\\src"
```

**Erreur "Python not found"**
```
Vérifie que Python est dans le PATH système
Ou spécifie le chemin complet dans la config:
"command": "C:\\Python311\\python.exe"
```

**Erreur JSON syntax**
```
Utilise un validateur JSON en ligne pour vérifier config.json
Vérifie les virgules, guillemets, accolades
```

---

### Le serveur se connecte mais les nouveaux outils n'apparaissent pas

**Cause:** Cache ou ancienne version du serveur chargée

**Solution:**
1. Ferme Claude Code **complètement**
2. Supprime le cache (optionnel):
   ```
   Supprimer: %USERPROFILE%\.claude\cache
   ```
3. Redémarre Claude Code
4. Vérifie avec "What VBA MCP tools do you have available?"

---

### Les anciens outils fonctionnent mais pas les nouveaux

**Cause:** Erreur de syntaxe dans le nouveau code

**Test manuel du serveur:**

1. Ouvre **Invite de commandes** (CMD)
2. Navigue vers le projet:
   ```
   cd C:\Users\alexi\Documents\projects\vba-mcp-monorepo
   ```
3. Exécute le test:
   ```
   test_server_standalone.bat
   ```

**Résultat attendu:**
```
[1/3] Setting environment...
[2/3] Testing server import...
[OK] Server imported successfully

[3/3] Listing MCP tools...
[OK] Server has 21 tools:
  - extract_vba
  - list_modules
  - analyze_code
  - inject_vba
  - validate_vba_code
  - list_macros
  - open_in_office
  - run_macro
  - get_worksheet_data
  - set_worksheet_data
  - close_office_file
  - list_open_files
  - list_tables         ← NOUVEAU
  - insert_rows         ← NOUVEAU
  - delete_rows         ← NOUVEAU
  - insert_columns      ← NOUVEAU
  - delete_columns      ← NOUVEAU
  - create_table        ← NOUVEAU
  - backup
  - refactor

[SUCCESS] Server is working!
```

Si tu vois des erreurs ici, c'est un problème de code Python, pas de configuration Claude Code.

---

## 📊 Checklist de Redémarrage

- [ ] 1. Fermer Claude Code complètement (vérifier system tray)
- [ ] 2. Attendre 5 secondes
- [ ] 3. Relancer Claude Code
- [ ] 4. Attendre l'icône marteau 🔨 (5-10 secondes)
- [ ] 5. Tester: "What VBA MCP tools do you have available?"
- [ ] 6. Vérifier que 21 outils sont listés
- [ ] 7. Vérifier que les 6 nouveaux outils Excel Tables apparaissent
- [ ] 8. Tester un outil Excel Tables (list_tables ou create_table)

**Temps total:** ~30 secondes

---

## 🎯 Tests Rapides Post-Redémarrage

### Test Excel Tables Basique

```
In test.xlsm:
1. Create an Excel table named "QuickTest" from range A1:C5 on Sheet1
2. List all tables
3. Insert a column named "Status" at position D in QuickTest
4. Get the table data to verify
```

**Si ça marche:** Tout est bon ! 🎉

**Si ça échoue:** Consulte la section Dépannage ci-dessus.

---

## 📚 Prochaines Étapes

Après un redémarrage réussi:

1. **Consulte** `QUICK_TEST_PROMPTS.md` pour tester tous les nouveaux outils
2. **Essaye** les prompts Excel Tables dans `PROMPTS_READY_TO_USE.md`
3. **Lis** `CHANGELOG.md` pour voir tous les détails de la v0.3.0

---

## 💡 Conseil Pro

**Créer un raccourci de redémarrage rapide:**

1. Crée un fichier `restart_claude.bat`:
   ```batch
   @echo off
   echo Fermeture de Claude Code...
   taskkill /F /IM "Claude.exe" 2>nul
   timeout /t 3 /nobreak >nul
   echo Redémarrage de Claude Code...
   start "" "C:\Users\alexi\AppData\Local\Programs\Claude\Claude.exe"
   echo Done!
   ```

2. Double-clic pour redémarrer instantanément !

---

**Version:** v0.3.0
**Date:** 2025-12-14
**Nouveautés:** 6 outils Excel Tables, 21 outils total
