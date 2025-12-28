# Quick Start - HTTP/SSE Transport

## TL;DR

```bash
# Windows - Démarrer le serveur
start_http_server.bat

# WSL/Mac/Linux - Se connecter
# Config: http://WINDOWS_IP:8000/sse
```

## Étape 1: Démarrer le serveur (Windows)

### Option A: Double-clic
```
Double-clic sur: start_http_server.bat
```

### Option B: Ligne de commande
```bash
venv\Scripts\vba-mcp-server-pro-http.exe --host 0.0.0.0 --port 8000
```

Le serveur affichera:
```
Starting server on 0.0.0.0:8000
SSE endpoint: http://0.0.0.0:8000/sse
```

## Étape 2: Trouver l'IP Windows

### Depuis Windows
```cmd
ipconfig
```
Chercher: `IPv4 Address.........: 192.168.x.x`

### Depuis WSL
```bash
cat /etc/resolv.conf | grep nameserver | awk '{print $2}'
# Output: 172.28.160.1
```

## Étape 3: Configurer le client

### Claude Code

**Fichier**:
- macOS: `~/.claude/config.json`
- Windows: `%USERPROFILE%\.claude\config.json`
- WSL: `~/.claude/config.json`

**Contenu**:
```json
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://172.28.160.1:8000/sse"
    }
  }
}
```

### Claude Code

```bash
# Créer/éditer ~/.claude/config.json
cat > ~/.claude/config.json <<EOF
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://172.28.160.1:8000/sse"
    }
  }
}
EOF
```

## Étape 4: Tester

Redémarrer Claude Code, puis:

```
"List all VBA modules in C:\Users\alexi\Documents\test.xlsm"
```

## Troubleshooting

### Connexion refusée
- Vérifier que le serveur tourne sur Windows
- Vérifier l'IP dans la config
- Vérifier le firewall Windows

### Serveur ne démarre pas
- Vérifier que le port 8000 est libre
- Essayer un autre port: `--port 8001`

### Outils ne fonctionnent pas
- Utiliser des chemins Windows: `C:\Users\...`
- PAS de chemins WSL: `/mnt/c/...`

## Scripts rapides

### Obtenir l'IP Windows depuis WSL
```bash
WIN_IP=$(cat /etc/resolv.conf | grep nameserver | awk '{print $2}')
echo "Windows IP: $WIN_IP"
```

### Configuration automatique depuis WSL
```bash
WIN_IP=$(cat /etc/resolv.conf | grep nameserver | awk '{print $2}')

cat > ~/.claude/config.json <<EOF
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://${WIN_IP}:8000/sse"
    }
  }
}
EOF

echo "Configuration créée avec IP: $WIN_IP"
```

## Voir aussi

- Guide complet: [packages/pro/HTTP_SETUP.md](packages/pro/HTTP_SETUP.md)
- Configuration exemple: [claude_desktop_config_http.json](claude_desktop_config_http.json) (exemple pour référence)
