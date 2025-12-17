# Push to GitHub

## Option 1: Créer le repo sur GitHub.com

1. Va sur https://github.com/new
2. Nom: `vba-mcp-monorepo`
3. **IMPORTANT**: Choisis **Private** (contient le code Pro)
4. Ne coche PAS "Initialize with README"
5. Crée le repo

Puis exécute:

```bash
cd C:\Users\alexi\Documents\projects\vba-mcp-monorepo
git remote add origin git@github.com:AlexisTrouve/vba-mcp-monorepo.git
git push -u origin main
```

## Option 2: Installer GitHub CLI

```bash
winget install GitHub.cli
gh auth login
gh repo create AlexisTrouve/vba-mcp-monorepo --private --source=. --push
```

## Rappel: Pour publier la version publique (sans Pro)

```bash
# Copie le .gitignore public
cp .gitignore.public .gitignore

# Crée un repo public séparé ou push sur vba-mcp-server existant
git remote add public git@github.com:AlexisTrouve/vba-mcp-server.git
git push public main
```
