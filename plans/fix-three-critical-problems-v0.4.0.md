# Plan d'Implémentation : Fix 3 Problèmes Critiques VBA MCP Pro

**Version:** 0.4.0
**Estimé:** 16-20 heures sur 3-4 jours
**Risque:** Moyen (nécessite tests approfondis avec Office)

---

## 🎯 Objectifs

Résoudre les 3 problèmes critiques identifiés dans RETRY_RESULTS.md:

1. **inject_vba** - Faux positifs (dit "Success" mais module n'existe pas)
2. **run_macro** - Macros injectées inexécutables (bloquées par sécurité Excel)
3. **Corruption fichiers** - Injections multiples corrompent les fichiers

**Score actuel:** 6/10
**Score cible:** 9/10 (production-ready)

---

## 📋 PHASE 1: Infrastructure (4-5h)

### Objectif
Éliminer corruption en unifiant l'accès aux fichiers via session_manager.

### 1.1 Refactorer inject_vba → Utiliser session_manager

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/inject.py`

**Problème actuel:**
- Ligne 312: `app = win32com.client.Dispatch(app_name)` crée instance séparée
- Bypass session_manager → accès concurrent → corruption

**Changements:**

**1. Modifier `inject_vba_tool()` (lignes 157-269):**
```python
# AVANT (ligne 237-244):
if platform.system() != "Windows":
    raise NotImplementedError(...)

return _inject_vba_windows(...)  # ← Crée nouvelle instance COM

# APRÈS:
if platform.system() != "Windows":
    raise NotImplementedError(...)

# Utiliser session_manager au lieu de créer nouvelle instance
from vba_mcp_pro.session_manager import OfficeSessionManager
manager = OfficeSessionManager.get_instance()
session = await manager.get_or_create_session(path)

return await _inject_vba_via_session(session, module_name, code, backup_path)
```

**2. Remplacer `_inject_vba_windows()` par `_inject_vba_via_session()`:**

Supprimer lignes 272-430 (fonction actuelle).

Ajouter nouvelle fonction:
```python
async def _inject_vba_via_session(
    session: OfficeSession,
    module_name: str,
    code: str,
    backup_path: Optional[Path] = None
) -> dict:
    """
    Inject VBA via OfficeSession existante (évite accès concurrent).

    Utilise session.app, session.file_obj, session.vb_project.
    Pas de CoInitialize/CoUninitialize (géré par session).
    """
    # Logique identique à _inject_vba_windows SAUF:
    # - Pas de création app/file_obj (utilise session)
    # - Pas de CoInitialize/CoUninitialize
    # - Pas de app.Quit() (session reste ouverte)
```

**Trade-offs:**
- ✅ Élimine accès concurrent → plus de corruption
- ✅ Réutilise sessions → injection plus rapide
- ❌ Changement architectural majeur
- ⚠️ Dépendance session_manager (doit être robuste)

---

### 1.2 Améliorer COM Cleanup

**Fichier:** `packages/pro/src/vba_mcp_pro/session_manager.py`

**Problème actuel:**
- Ligne 387: `session.app.Quit()` sans `pythoncom.ReleaseObject()`
- Références COM persistent → locks fichiers

**Changements:**

**1. Ajouter helper (après ligne 506):**
```python
def _release_com_objects(self, session: OfficeSession) -> None:
    """Release explicite objets COM."""
    try:
        import pythoncom

        if session._vb_project is not None:
            pythoncom.ReleaseObject(session._vb_project)
            session._vb_project = None

        if session.file_obj is not None:
            pythoncom.ReleaseObject(session.file_obj)

        if session.app is not None:
            pythoncom.ReleaseObject(session.app)

    except Exception as e:
        logger.warning(f"Error releasing COM: {e}")
```

**2. Modifier `_close_session_internal()` (lignes 352-398):**
```python
# AVANT ligne 390 (dans finally):
try:
    pythoncom.CoUninitialize()
except Exception as e:
    logger.warning(f"Failed to uninitialize COM: {e}")

# APRÈS:
# D'abord release objets COM
self._release_com_objects(session)

# Puis CoUninitialize
try:
    pythoncom.CoUninitialize()
except Exception as e:
    logger.warning(f"Failed to uninitialize COM: {e}")
```

**Trade-offs:**
- ✅ Élimine locks fichiers
- ✅ Prévient memory leaks
- ❌ Cleanup plus complexe
- ⚠️ Release agressif peut causer erreurs si session référencée ailleurs

---

### 1.3 Ajouter File Lock Detection

**Fichier:** `packages/pro/src/vba_mcp_pro/session_manager.py`

**Ajouter helper (après `_release_com_objects`):**
```python
def _check_file_lock(self, file_path: Path) -> bool:
    """Vérifie si fichier verrouillé par autre processus."""
    import win32file
    import pywintypes

    try:
        handle = win32file.CreateFile(
            str(file_path),
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0,  # Pas de partage
            None,
            win32file.OPEN_EXISTING,
            0,
            None
        )
        win32file.CloseHandle(handle)
        return False  # Pas verrouillé
    except pywintypes.error:
        return True  # Verrouillé
```

**Modifier `get_or_create_session()` (lignes 145-189):**
```python
# Avant d'ouvrir fichier, vérifier lock
if self._check_file_lock(file_path):
    # Si dans nos sessions, vérifier si vivante
    if file_key in self._sessions:
        session = self._sessions[file_key]
        if session.is_alive():
            session.refresh_last_accessed()
            return session
        else:
            # Session morte mais fichier locked → processus externe
            raise PermissionError(
                f"File is locked by another process: {file_path}"
            )
    else:
        # Pas dans nos sessions → locked par externe
        raise PermissionError(
            f"File is locked by another application. "
            f"Close the file in Excel/Word and try again."
        )
```

**Trade-offs:**
- ✅ Messages erreur clairs
- ✅ Prévient corruption accès concurrent
- ❌ I/O supplémentaire à chaque session
- ⚠️ Faux positifs possibles sur réseaux

---

## 📋 PHASE 2: Validation (3-4h)

### Objectif
Éliminer faux positifs en vérifiant que l'injection a réellement fonctionné.

### 2.1 Ajouter Post-Save Verification

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/inject.py`

**Ajouter fonction (après `_compile_vba_module`, ligne 156):**
```python
async def _verify_injection(
    file_path: Path,
    module_name: str,
    expected_code: str
) -> Tuple[bool, Optional[str]]:
    """
    Vérifie injection en rouvrant le fichier.
    Détecte cas où Save() réussit mais module non persisté.
    """
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    app = None
    file_obj = None

    try:
        # Ouvrir en read-only
        app = win32com.client.Dispatch("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False

        file_obj = app.Workbooks.Open(str(file_path), ReadOnly=True)
        vb_project = file_obj.VBProject

        # Chercher module
        for component in vb_project.VBComponents:
            if component.Name == module_name:
                code_module = component.CodeModule
                if code_module.CountOfLines > 0:
                    actual_code = code_module.Lines(1, code_module.CountOfLines)

                    # Comparer
                    if actual_code.strip() != expected_code.strip():
                        return False, "Code mismatch in saved file"
                else:
                    return False, "Module exists but is empty"
                break
        else:
            return False, f"Module '{module_name}' not found in saved file"

        return True, None

    except Exception as e:
        return False, f"Verification failed: {str(e)}"

    finally:
        if file_obj:
            try:
                file_obj.Close(SaveChanges=False)
            except:
                pass
        if app:
            try:
                app.Quit()
            except:
                pass
        pythoncom.CoUninitialize()
```

**Modifier `_inject_vba_via_session()` pour appeler verification:**
```python
# Après save (équivalent ligne 412-418 actuelle):
if "Excel" in session.app_type:
    session.file_obj.Save()
elif "Word" in session.app_type:
    session.file_obj.Save()

# AJOUTER VERIFICATION:
success, error = await _verify_injection(path, module_name, code)
if not success:
    # Restaurer backup si existe
    if backup_path and backup_path.exists():
        shutil.copy2(backup_path, path)
        raise ValueError(
            f"Injection verification failed: {error}\n"
            f"File restored from backup: {backup_path}"
        )
    else:
        raise ValueError(f"Injection verification failed: {error}")

# Si OK, continuer
return {
    "action": action,
    "module": module_name,
    "validated": True,
    "verified": True  # NOUVEAU
}
```

**Trade-offs:**
- ✅ Détecte faux positifs (problème #1)
- ✅ Auto-rollback si échec
- ❌ Injection plus lente (re-open fichier)
- ⚠️ Verification elle-même peut échouer (timing)

---

### 2.2 Améliorer _compile_vba_module

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/inject.py`

**Problème actuel (lignes 103-155):**
- Lecture lignes ne force PAS compilation
- Ligne 147-154: Exception → retourne `True` (masque erreurs!)

**Remplacer fonction complète:**
```python
def _compile_vba_module(vb_module) -> Tuple[bool, Optional[str]]:
    """
    Valide module en forçant parsing VBA.
    Utilise ProcOfLine() qui force analyse sémantique.
    """
    try:
        import pythoncom

        code_module = vb_module.CodeModule
        line_count = code_module.CountOfLines

        if line_count == 0:
            return True, None

        # Lire tout le code (force parsing)
        try:
            full_code = code_module.Lines(1, line_count)
        except pythoncom.com_error as e:
            return False, f"Failed to read code: {str(e)}"

        # Accéder ProcOfLine pour chaque ligne (force semantic checks)
        for line_num in range(1, min(line_count + 1, 1000)):
            try:
                # ProcOfLine lève exception si syntax error
                proc_name = code_module.ProcOfLine(line_num, 0)
            except pythoncom.com_error as e:
                error_msg = str(e)
                if "Compile error" in error_msg or "Syntax error" in error_msg:
                    return False, f"Syntax error at line {line_num}: {error_msg}"
                # Autres erreurs OK (ligne hors procédure)

        # Vérifier propriétés basiques
        try:
            _ = vb_module.Name
            _ = code_module.CountOfDeclarationLines
        except pythoncom.com_error as e:
            return False, f"Module validation error: {str(e)}"

        return True, None

    except Exception as e:
        # NE PLUS MASQUER - propager erreurs inattendues
        error_str = str(e)
        if "Compile error" in error_str or "Syntax error" in error_str:
            return False, f"Validation error: {error_str}"
        # Pour autres erreurs, on lève exception
        raise
```

**Trade-offs:**
- ✅ Détection erreurs syntaxe améliorée
- ✅ Plus de masquage silencieux
- ❌ Plus lent (itère toutes lignes)
- ⚠️ Limité par API COM (certaines erreurs non détectables)

---

### 2.3 Supprimer Exception Masking

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/inject.py`

**Changements:**

**1. Ligne 364 (bare except):**
```python
# AVANT:
except:
    vb_component = vb_project.VBComponents.Add(1)

# APRÈS:
except AttributeError:  # Constante non disponible
    vb_component = vb_project.VBComponents.Add(1)
except Exception as e:
    logger.error(f"Failed to add component: {e}")
    raise
```

**2. Ligne 427 (bare except):**
```python
# AVANT:
except:
    pass

# APRÈS:
except Exception as e:
    logger.warning(f"Cleanup error: {e}")
    # Ne pas masquer - mais continuer cleanup
```

**3. Lignes 403-408 (masquage PermissionError):**
```python
# AVANT:
except Exception as e:
    raise PermissionError(
        f"Cannot access VBA project. "
        f"Ensure 'Trust access...' is enabled"
    ) from e

# APRÈS:
except pythoncom.com_error as e:
    error_msg = str(e).lower()
    if "permission" in error_msg or "access denied" in error_msg:
        raise PermissionError(
            f"Cannot access VBA project. "
            f"Ensure 'Trust access...' is enabled"
        ) from e
    # Autres erreurs COM - propager avec type original
    raise RuntimeError(f"COM error during injection: {str(e)}") from e
except Exception as e:
    # Erreurs non-COM - propager avec type original
    raise
```

**Trade-offs:**
- ✅ Erreurs claires pour debugging
- ✅ Types exception corrects
- ❌ Plus d'exceptions exposées aux users
- ⚠️ Peut casser code qui assume PermissionError

---

## 📋 PHASE 3: Sécurité (2-3h)

### Objectif
Permettre exécution macros injectées en modifiant temporairement AutomationSecurity.

### 3.1 Ajouter AutomationSecurity Context Manager

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/office_automation.py`

**Ajouter classe (après imports, ligne 18):**
```python
class AutomationSecurityContext:
    """
    Context manager pour abaisser temporairement AutomationSecurity Office.
    Permet exécution macros sans prompts utilisateur.
    Restaure niveau original automatiquement.
    """

    def __init__(self, app, target_level: int = 1):
        """
        Args:
            app: Application COM Office
            target_level:
                1 = msoAutomationSecurityLow (macros enabled)
                2 = msoAutomationSecurityByUI (user setting)
                3 = msoAutomationSecurityForceDisable (disabled)
        """
        self.app = app
        self.target_level = target_level
        self.original_level = None

    def __enter__(self):
        """Abaisser sécurité."""
        try:
            self.original_level = self.app.AutomationSecurity
            logger.info(f"AutomationSecurity: {self.original_level} → {self.target_level}")
            self.app.AutomationSecurity = self.target_level
        except Exception as e:
            logger.warning(f"Cannot modify AutomationSecurity: {e}")
            # Continuer quand même
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Restaurer sécurité."""
        if self.original_level is not None:
            try:
                self.app.AutomationSecurity = self.original_level
                logger.info(f"AutomationSecurity restored: {self.original_level}")
            except Exception as e:
                logger.warning(f"Cannot restore AutomationSecurity: {e}")
        return False
```

---

### 3.2 Modifier run_macro_tool

**Fichier:** `packages/pro/src/vba_mcp_pro/tools/office_automation.py`

**Modifier signature (ligne 150):**
```python
async def run_macro_tool(
    file_path: str,
    macro_name: str,
    arguments: Optional[List[Any]] = None,
    enable_macros: bool = True  # NOUVEAU paramètre
) -> str:
```

**Modifier exécution (autour ligne 220):**
```python
# AVANT:
last_error = None
for format_name in formats_to_try:
    try:
        result = session.app.Run(format_name, *args)
        # ...

# APRÈS:
last_error = None

# Context manager sécurité
if enable_macros:
    security_context = AutomationSecurityContext(session.app, target_level=1)
else:
    from contextlib import nullcontext
    security_context = nullcontext()

with security_context:
    for format_name in formats_to_try:
        try:
            result = session.app.Run(format_name, *args)
            # ... (reste identique)
```

**Mettre à jour MCP schema (dans server.py):**
```python
# Ajouter dans schema de run_macro:
{
    "name": "enable_macros",
    "type": "boolean",
    "description": "Enable macros by temporarily lowering AutomationSecurity (default: true)",
    "default": True
}
```

**Trade-offs:**
- ✅ Macros injectées exécutables (problème #2 résolu!)
- ✅ Sécurité restaurée automatiquement
- ❌ Abaisse sécurité temporairement
- ⚠️ Macros malicieuses pourraient s'exécuter (mais backup existe)

---

## 📋 PHASE 4: Tests (3-4h)

### 4.1 Tests Unitaires

**Fichiers:**
- `packages/pro/tests/test_inject.py`
- `packages/pro/tests/test_office_automation_tools.py`
- `packages/pro/tests/test_session_manager.py`

**Nouveaux tests à ajouter:**

**test_inject.py:**
1. `test_inject_uses_session_manager()` - Mock session, vérifier pas de Dispatch direct
2. `test_inject_detects_concurrent_access()` - Mock file lock, assert PermissionError
3. `test_inject_post_save_verification()` - Mock verification failure, assert rollback
4. `test_inject_no_exception_masking()` - Vérifier types exception corrects

**test_office_automation_tools.py:**
1. `test_run_macro_with_security_context()` - Mock AutomationSecurity, vérifier abaissé/restauré
2. `test_run_macro_security_restored_on_error()` - Macro échoue, security quand même restaurée
3. `test_run_macro_without_security_modification()` - enable_macros=False, pas de modif

**test_session_manager.py:**
1. `test_com_cleanup_releases_objects()` - Mock ReleaseObject, vérifier appelé
2. `test_file_lock_detection()` - Mock lock, assert PermissionError

---

### 4.2 Tests d'Intégration

**Nouveau fichier:** `packages/pro/tests/test_integration_full_workflow.py`

**Tests critiques:**

```python
@pytest.mark.integration
@pytest.mark.windows_only
async def test_inject_and_run_macro_full_workflow():
    """Test complet: inject → verify → run → vérifier pas corruption"""
    # 1. Injecter macro test
    # 2. Vérifier module existe (post-save verification)
    # 3. Exécuter macro
    # 4. Fermer session
    # 5. Rouvrir fichier, vérifier module toujours là

@pytest.mark.integration
async def test_multiple_injections_no_corruption():
    """Test 10 injections consécutives, vérifier pas corruption"""
    # 1. Boucle 10 injections différents modules
    # 2. Vérifier tous modules existent
    # 3. Fichier ouvrable dans Excel
    # 4. Pas de processus zombie
```

**Fixtures nécessaires:**
```python
@pytest.fixture
def sample_xlsm_copy(tmp_path):
    """Copie test.xlsm dans répertoire temporaire"""
    # ...
```

---

### 4.3 Guide Test Manuel

**Nouveau fichier:** `packages/pro/MANUAL_TEST_CHECKLIST.md`

**Checklist essentielle:**

```markdown
## Test 1: Injection via Session Manager
1. Ouvrir fichier avec open_in_office
2. Injecter module
3. Vérifier pas d'erreur "file locked"
4. Injecter 2e module
5. Ouvrir Excel manuellement, vérifier 2 modules présents

## Test 2: Post-Save Verification
1. Injecter code valide
2. Vérifier message "Verified: Yes"
3. Rouvrir fichier, module existe

## Test 3: Macro Execution avec Security
1. Injecter macro
2. run_macro avec enable_macros=true
3. Macro s'exécute sans prompt
4. Vérifier logs: AutomationSecurity abaissé puis restauré

## Test 4: Multiple Injections
1. 10 injections consécutives
2. Pas de corruption
3. Tous modules présents
4. Task Manager: pas de processus zombie Excel

## Test 5: Error Handling
1. Code avec syntax error → erreur claire (pas "Success")
2. Code avec Unicode → suggestions ASCII
3. Fichier locked → message clair
```

---

## 📋 PHASE 5: Documentation (1-2h)

### 5.1 Mise à jour fichiers

**KNOWN_ISSUES.md:**
- Ajouter sections "Issue #5, #6, #7 RESOLVED"
- Détailler solutions implémentées
- Enlever de "Known Issues" (déplacer vers "Resolved")

**CHANGELOG.md:**
- Créer section `[0.4.0] - 2025-12-15`
- Lister tous changements (Fixed, Added, Changed, Removed)
- Migration guide si breaking changes

**README.md (Pro):**
- Mettre à jour "Known Limitations"
- Ajouter note sur AutomationSecurity
- Exemples avec enable_macros parameter

**Inline comments:**
- Docstrings pour nouvelles fonctions
- Comments critiques (security, rollback logic)

---

## 📂 Fichiers Critiques

| Fichier | Lignes Modifiées | Changements |
|---------|-----------------|-------------|
| `inject.py` | 157-269, 272-430 | Refactor session_manager, verification, validation |
| `session_manager.py` | 352-398, +506 | COM cleanup, file lock detection |
| `office_automation.py` | +18, 150-268 | AutomationSecurity context manager |
| `test_integration_full_workflow.py` | NEW | Tests end-to-end |
| `MANUAL_TEST_CHECKLIST.md` | NEW | Guide test manuel |
| `KNOWN_ISSUES.md` | +sections | Documenter résolutions |
| `CHANGELOG.md` | +v0.4.0 | Changelog complet |

---

## ⚠️ Risques & Mitigations

| Risque | Impact | Probabilité | Mitigation |
|--------|--------|-------------|------------|
| ReleaseObject() cause crashes | HIGH | LOW | Tests unitaires, null checks |
| AutomationSecurity pas restaurée | HIGH | LOW | Context manager `__exit__` garanti |
| Verification bloque injections valides | MEDIUM | MEDIUM | Try-catch avec fallback |
| Session manager cassé | MEDIUM | LOW | Tests unitaires complets |
| Exceptions inattendues | LOW | MEDIUM | Exception mapping, logging |

---

## ✅ Critères de Succès

**Phase 1:**
- [ ] inject_vba utilise OfficeSessionManager
- [ ] Pas d'instance COM séparée créée
- [ ] COM objects released explicitement
- [ ] File lock détecté et bloqué

**Phase 2:**
- [ ] Post-save verification fonctionne
- [ ] Faux positifs détectés et bloqués
- [ ] _compile_vba_module détecte erreurs syntaxe
- [ ] Exceptions non masquées

**Phase 3:**
- [ ] run_macro exécute macros injectées
- [ ] AutomationSecurity restaurée (même si erreur)
- [ ] enable_macros parameter fonctionne
- [ ] Logging sécurité complet

**Phase 4:**
- [ ] Tous tests unitaires passent
- [ ] Tests intégration passent (Windows + Office)
- [ ] 10 injections consécutives sans corruption
- [ ] Pas de processus zombie

**Phase 5:**
- [ ] KNOWN_ISSUES.md à jour
- [ ] CHANGELOG.md v0.4.0 créé
- [ ] Guide test manuel validé
- [ ] Documentation technique complète

---

## 🚀 Ordre d'Implémentation

**Jour 1 (6-8h):**
1. Phase 1.1: Refactor inject_vba
2. Phase 1.2: COM cleanup
3. Phase 1.3: File lock detection
4. Tests unitaires Phase 1

**Jour 2 (5-6h):**
1. Phase 2.1: Post-save verification
2. Phase 2.2: Fix _compile_vba_module
3. Phase 2.3: Remove exception masking
4. Tests unitaires Phase 2

**Jour 3 (4-5h):**
1. Phase 3.1: AutomationSecurity
2. Phase 3.2: Modifier run_macro
3. Tests unitaires Phase 3
4. Tests intégration complets

**Jour 4 (2-3h):**
1. Tests manuels (checklist complète)
2. Documentation (CHANGELOG, KNOWN_ISSUES)
3. Inline comments
4. Review final

---

## 🔄 Rollback Strategy

**Git:**
- Branch: `fix/three-critical-problems`
- Commits atomiques par phase
- Si échec: `git revert` ou `git reset --hard`

**Backups:**
- Système backup inject_vba déjà en place
- `.vba_backups/` folder

**Rollback incrémental:**
- Chaque phase indépendante
- Possible de rollback Phase 3 mais garder Phase 1+2

---

## 📊 Estimation Timeline

**Optimiste:** 12h (tout fonctionne du 1er coup)
**Réaliste:** 16h (debugging COM, tests)
**Pessimiste:** 24h (problèmes COM majeurs, corruption tests)

**Recommandé:** Planifier 16-20h sur 3-4 jours
