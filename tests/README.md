# Tests - VBA MCP Server

## Vue d'ensemble

Ce dossier contient les tests unitaires et d'intégration pour le VBA MCP Server.

## Structure

```
tests/
├── __init__.py                # Init du package tests
├── test_office_handler.py     # Tests pour OfficeHandler
├── test_vba_parser.py         # Tests pour VBAParser
└── test_tools.py              # Tests pour les tools MCP
```

## Lancer les tests

### Tous les tests

```bash
# Avec l'environnement virtuel activé
pytest tests/ -v

# Ou directement avec le venv
./venv/Scripts/pytest tests/ -v
```

### Tests spécifiques

```bash
# Tester uniquement le parser
pytest tests/test_vba_parser.py -v

# Tester uniquement l'office handler
pytest tests/test_office_handler.py -v

# Tester uniquement les tools
pytest tests/test_tools.py -v
```

### Avec couverture de code

```bash
pytest tests/ --cov=src --cov-report=html
```

Ouvre ensuite `htmlcov/index.html` dans un navigateur.

## Tests qui nécessitent un fichier Excel

Certains tests sont **skipped** par défaut car ils nécessitent un fichier Excel réel :

- `test_extract_real_excel_file`
- `test_extract_real_file`
- `test_extract_specific_module`
- `test_list_real_file`
- `test_analyze_real_file`
- `test_full_workflow`

### Pour activer ces tests :

1. **Créez le fichier Excel de test** :
   - Ouvrez Excel
   - Appuyez sur `Alt + F11` pour ouvrir l'éditeur VBA
   - Insertion → Module
   - Copiez le code VBA depuis `examples/sample_vba_code.txt`
   - Enregistrez sous `examples/test_simple.xlsm`

2. **Relancez les tests** :
   ```bash
   pytest tests/ -v
   ```

Les 6 tests précédemment skipped devraient maintenant s'exécuter !

## Statistiques actuelles

- **Total tests** : 29
- **Tests passés** : 23 ✅
- **Tests skipped** : 6 ⏭️ (nécessitent fichier Excel)
- **Tests échoués** : 0 ❌

## Écrire de nouveaux tests

### Exemple de test simple

```python
def test_my_feature(parser):
    \"\"\"Test my new feature\"\"\"
    result = parser.my_function()
    assert result == expected_value
```

### Exemple de test async

```python
@pytest.mark.asyncio
async def test_my_async_feature():
    \"\"\"Test async feature\"\"\"
    result = await my_async_function()
    assert result is not None
```

### Exemple de test conditionnel

```python
@pytest.mark.skipif(
    not Path("examples/test.xlsm").exists(),
    reason="Test file not available"
)
def test_with_real_file():
    \"\"\"Test that requires a real file\"\"\"
    # ...
```

## Dépendances de test

Les dépendances suivantes sont nécessaires (déjà dans `requirements.txt`) :

- `pytest>=7.4.0` - Framework de test
- `pytest-asyncio>=1.3.0` - Support tests async
- `pytest-cov>=4.1.0` - Couverture de code

## CI/CD

Les tests sont automatiquement exécutés via GitHub Actions sur chaque push/PR.

Voir `.github/workflows/tests.yml` pour la configuration.

## Debugging

### Mode verbeux
```bash
pytest tests/ -vv
```

### Afficher les prints
```bash
pytest tests/ -s
```

### Arrêter au premier échec
```bash
pytest tests/ -x
```

### Lancer un test spécifique
```bash
pytest tests/test_vba_parser.py::TestVBAParser::test_parse_simple_sub -v
```

## Bonnes pratiques

1. **Nommage** : Les tests doivent commencer par `test_`
2. **Isolation** : Chaque test doit être indépendant
3. **Fixtures** : Utiliser des fixtures pour le code réutilisable
4. **Assertions** : Une assertion principale par test
5. **Documentation** : Chaque test doit avoir une docstring claire

## Questions fréquentes

### Pourquoi certains tests sont skipped ?

Ces tests nécessitent un vrai fichier Excel avec VBA, qui ne peut pas être généré automatiquement. Créez `examples/test_simple.xlsm` pour les activer.

### Comment créer des mocks de fichiers Office ?

Pour des tests unitaires purs, on peut mocker `OfficeHandler` pour éviter de dépendre de vrais fichiers.

### Les tests sont lents ?

Optimisations possibles :
- Utiliser des fixtures avec `scope="session"`
- Paralléliser avec `pytest-xdist`
- Limiter la taille des fichiers de test
