# CLAUDE.md

This file provides guidance to Claude Code when working with this repository.

## Projet

**VBA MCP Server** est un serveur MCP (Model Context Protocol) qui permet à Claude Code d'extraire et d'analyser du code VBA depuis des fichiers Microsoft Office (Excel, Access, Word).

## Stratégie Hybride (IMPORTANT)

Ce projet suit une **stratégie hybride** pour le freelancing :

### Version Lite (Open Source - GitHub Public)
- ✅ Extraction VBA read-only
- ✅ Liste des modules
- ✅ Analyse de structure et complexité
- ✅ Portfolio technique visible
- 📦 Repo : https://github.com/AlexisTrouve/vba-mcp-server

### Version Pro (Privée - Repo séparé)
- 🔒 Modification et réinjection de VBA
- 🔒 Refactoring automatisé avec IA
- 🔒 Exécution de macros (sandboxed)
- 🔒 Testing framework
- 🔒 Version control integration
- 💰 Monétisation : $49-199/mois

**⚠️ RÈGLE CRITIQUE** :
- Le code de modification/réinjection VBA **NE DOIT JAMAIS** être committé dans ce repo public
- Les features pro restent dans un repo privé séparé
- Ce repo sert de portfolio et d'outil de base utilisable

## État Actuel

**Version actuelle** : 1.0.0 (Lite - En développement)

### Complété ✅
- Structure du projet
- Documentation complète (README, API, ARCHITECTURE, EXAMPLES)
- Serveur MCP fonctionnel (stdio transport)
- 3 outils MCP implémentés :
  - `extract_vba` - Extraction de code VBA
  - `list_modules` - Liste des modules
  - `analyze_structure` - Analyse structure/complexité
- Librairies core :
  - `OfficeHandler` - Gestion fichiers Office
  - `VBAParser` - Parsing VBA

### À faire ⏳
- [ ] Tests unitaires (pytest)
- [ ] Exemple de fichiers Office avec VBA
- [ ] Test avec projet VBA réel
- [ ] CI/CD GitHub Actions
- [ ] Support .xlsb, .accdb, .docm
- [ ] Vidéo démo

## Structure du Projet

```
vba-mcp-server/
├── src/
│   ├── server.py              # Point d'entrée MCP
│   ├── tools/                 # Outils MCP
│   │   ├── extract.py         # Extraction VBA
│   │   ├── list_modules.py    # Liste modules
│   │   └── analyze.py         # Analyse structure
│   └── lib/                   # Logique métier
│       ├── office_handler.py  # Gestion fichiers Office
│       └── vba_parser.py      # Parser VBA
├── docs/
│   ├── API.md                 # Référence API des tools
│   ├── ARCHITECTURE.md        # Architecture technique
│   └── EXAMPLES.md            # Exemples d'usage
├── examples/                  # Fichiers Office exemples
├── tests/                     # Tests unitaires
├── README.md                  # Documentation principale
├── QUICKSTART.md              # Guide démarrage rapide
├── ROADMAP.md                 # Feuille de route
├── requirements.txt           # Dépendances Python
└── LICENSE                    # MIT (lite) + Commercial (pro)
```

## Technologies

### Stack
- **Python 3.8+**
- **MCP SDK** - Model Context Protocol
- **oletools** - Extraction VBA depuis OLE2/OOXML
- **openpyxl** - Parsing Excel
- **pywin32** (optionnel, Windows) - COM APIs Office

### Transport
- **stdio** (principal) - Pour usage local avec Claude Code
- **HTTP** (futur) - Pour usage remote

## Formats Supportés

| Format | Description | Status |
|--------|-------------|--------|
| `.xlsm` | Excel Macro-Enabled | ✅ Supporté |
| `.xlsb` | Excel Binary | 🚧 Planifié |
| `.accdb` | Access Database | 🚧 Planifié |
| `.docm` | Word Macro-Enabled | 🚧 Planifié |
| `.xls` | Legacy Excel | 🔮 Future |
| `.mdb` | Legacy Access | 🔮 Future |

## Principes de Développement

### Code Quality
- **PEP 8** pour le style Python
- **Type hints** partout où possible
- **Docstrings** pour toutes les fonctions publiques
- **Error handling** explicite et informatif

### Sécurité
- ⚠️ **JAMAIS exécuter** de macros VBA (version lite)
- Validation stricte des chemins de fichiers
- Limite de taille de fichier (100 MB)
- Messages d'erreur sécurisés (pas de stack traces)

### Performance
- Lazy loading des modules VBA
- Caching des résultats parsés
- Streaming pour gros fichiers
- Timeout appropriés

## Usage avec Claude Code

### Configuration
```json
{
  "mcpServers": {
    "vba": {
      "command": "python",
      "args": ["C:/path/to/vba-mcp-server/src/server.py"]
    }
  }
}
```

### Exemples de requêtes
```
"Extract VBA from budget.xlsm"
"List all modules in report.xlsm"
"Analyze the structure of my Excel file"
```

## Cas d'Usage Freelance

### Pourquoi VBA en 2025 ?
- ✅ Marché legacy massif (entreprises avec VBA en prod)
- ✅ Peu de concurrence (devs évitent VBA)
- ✅ Tarifs élevés (maintenance legacy)
- ✅ Besoin de modernisation/refactor

### Proposition de Valeur
1. **Vitesse** : Analyse/refactor 10x plus rapide avec IA
2. **Qualité** : Détection automatique de code smell
3. **Unique** : Outil propriétaire = différenciateur
4. **Premium** : Justifie tarifs plus élevés

### ROI Estimé
- **Investissement** : 1-2 semaines dev
- **Retour** : 1 mission VBA gagnée = rentabilisé
- **Timeline** : Court/moyen terme (2-5 ans avant obsolescence VBA)

## Features Pro (NE PAS IMPLÉMENTER ICI)

Ces features restent dans un repo privé :

### Version 2.0 (Pro)
- Modification de code VBA
- Réinjection dans fichiers Office
- Backup automatique avant modification
- Rollback de changements
- Exécution de macros (sandboxed)
- Refactoring automatisé avec IA

### Version 3.0 (Enterprise)
- Migration Access → Excel
- Conversion VBA → Python
- Collaboration multi-utilisateurs
- Dashboard web
- API REST
- Webhooks

## Testing

### Test avec fichier Excel
```python
# Créer un fichier Excel avec VBA simple
# Module1:
Sub HelloWorld()
    MsgBox "Hello from VBA!"
End Sub

# Tester l'extraction
python src/server.py --test examples/test.xlsm
```

### Tests unitaires
```bash
pytest tests/ -v
pytest tests/test_extract.py::test_extract_xlsm
```

## Publication GitHub

### Avant le push
- ✅ Vérifier aucun code pro inclus
- ✅ README complet et professionnel
- ✅ LICENSE correct (MIT pour lite)
- ✅ .gitignore approprié
- ✅ Documentation à jour

### Workflow Git
```bash
git init
git add .
git commit -m "Initial release: VBA MCP Server v1.0"
git remote add origin git@github.com:AlexisTrouve/vba-mcp-server.git
git push -u origin main
```

### GitHub Settings
- Description : "MCP server for VBA extraction and analysis from Office files"
- Topics : `mcp`, `vba`, `excel`, `office`, `claude-code`, `code-analysis`
- License : MIT
- README preview actif

## Marketing & Visibilité

### Contenu à créer
1. **Vidéo démo** (3-5 min) sur YouTube
2. **Article blog** sur Medium/DEV.to
3. **Post LinkedIn** avec démo
4. **Tweet** avec GIF de demo

### Pitch
> "Tired of manually analyzing VBA code? VBA MCP Server lets Claude Code extract, analyze, and help refactor your Office macros automatically. Open source, MIT licensed. Pro version available for enterprise."

## Métriques de Succès

### Version Lite (6 mois)
- 🎯 100+ GitHub stars
- 🎯 500+ installations
- 🎯 10+ contributors
- 🎯 Featured in MCP registry

### Version Pro (12 mois)
- 🎯 10+ clients payants
- 🎯 $1,000+ MRR
- 🎯 <10% churn
- 🎯 50+ NPS

## Maintenance

### Dépendances à surveiller
- **MCP SDK** - Mises à jour du protocole
- **oletools** - Nouvelles versions Office
- **Microsoft Office** - Changements de format

### Compatibilité
- Python 3.8 minimum (pour type hints)
- Windows, macOS, Linux
- Office 2007+ (OOXML)

## Ressources

### Documentation Externe
- [MCP Specification](https://modelcontextprotocol.io)
- [oletools Documentation](https://github.com/decalage2/oletools)
- [MS-OVBA Spec](https://docs.microsoft.com/en-us/openspecs/office_file_formats/)

### Communauté
- MCP Discord
- r/vba Reddit
- Stack Overflow (tag: vba)

## Notes Importantes

1. **Ne jamais** publier de clés API ou credentials
2. **Toujours** tester avec fichiers Office réels avant release
3. **Documenter** chaque changement dans CHANGELOG.md
4. **Versionner** selon SemVer (X.Y.Z)
5. **Séparer** strictement lite et pro (repos différents)

## Contact

- **Développeur** : Alexis Trouve
- **Email** : alexistrouve.pro@gmail.com
- **GitHub** : @AlexisTrouve
- **LinkedIn** : /in/alexistrouve

---

**Dernière mise à jour** : 2025-12-11
**Version du projet** : 1.0.0 (Lite - Dev)
