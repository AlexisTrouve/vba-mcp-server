# Claude Code + HTTP/SSE Setup Guide

Guide pour connecter Claude Code (depuis WSL/Mac/Linux) au serveur VBA MCP Pro (sur Windows).

## Prérequis

- **Serveur Windows**: VBA MCP Pro installé avec pywin32
- **Client**: Claude Code installé (peut être sur WSL, Mac, Linux)

## Étape 1: Démarrer le serveur (Windows)

### Option A: PowerShell (Recommandé pour test)

```powershell
cd C:\Users\alexi\Documents\projects\vba-mcp-monorepo

# Tester d'abord
.\test_http_windows.ps1

# Si OK, démarrer le serveur
.\start_http_server.bat
```

### Option B: Ligne de commande

```powershell
venv\Scripts\vba-mcp-server-pro-http.exe --host 0.0.0.0 --port 8000
```

Le serveur affichera:
```
SSE endpoint: http://0.0.0.0:8000/sse
Messages endpoint: http://0.0.0.0:8000/messages/
```

## Étape 2: Trouver l'IP Windows

### Depuis Windows
```powershell
ipconfig | findstr IPv4
# Example: IPv4 Address. . . . . . . . . : 192.168.1.100
```

### Depuis WSL
```bash
cat /etc/resolv.conf | grep nameserver | awk '{print $2}'
# Example: 172.28.160.1
```

## Étape 3: Configurer Claude Code

### Configuration manuelle

Créer/éditer `~/.config/claude-code/mcp_settings.json`:

```json
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://172.28.160.1:8000/sse"
    }
  }
}
```

**Note**: Remplacer `172.28.160.1` par l'IP Windows trouvée à l'étape 2.

### Script automatique (WSL/Linux)

```bash
#!/bin/bash
# Script: configure_claude_code.sh

# Détecter l'IP Windows
WIN_IP=$(cat /etc/resolv.conf | grep nameserver | awk '{print $2}')
echo "Detected Windows IP: $WIN_IP"

# Créer le répertoire de config
mkdir -p ~/.config/claude-code

# Créer la configuration
cat > ~/.config/claude-code/mcp_settings.json <<EOF
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://${WIN_IP}:8000/sse"
    }
  }
}
EOF

echo "✓ Configuration created at ~/.config/claude-code/mcp_settings.json"
echo "✓ Server URL: http://${WIN_IP}:8000/sse"
echo ""
echo "Next steps:"
echo "1. Make sure VBA MCP server is running on Windows"
echo "2. Test connection: claude --list-tools"
```

Rendre exécutable et lancer:
```bash
chmod +x configure_claude_code.sh
./configure_claude_code.sh
```

## Étape 4: Tester la connexion

### Test simple

```bash
# Depuis WSL/Mac/Linux
claude --list-tools
```

Vous devriez voir 21 tools listés.

### Test avec un fichier

```bash
# NOTE: Utiliser des chemins Windows (pas WSL)
claude "List all VBA modules in C:\Users\alexi\Documents\test.xlsm"
```

**IMPORTANT**:
- ✅ Utiliser: `C:\Users\...\file.xlsm`
- ❌ NE PAS utiliser: `/mnt/c/Users/.../file.xlsm`

Le serveur tourne sur Windows, donc il faut des chemins Windows!

## Troubleshooting

### Erreur: Connection refused

**Cause**: Le serveur n'est pas démarré ou firewall bloque.

**Solutions**:
1. Vérifier que le serveur tourne sur Windows
2. Vérifier l'IP est correcte
3. Désactiver temporairement le firewall Windows
4. Vérifier `--host 0.0.0.0` (pas `127.0.0.1`)

### Erreur: Tool execution failed

**Cause**: Chemin de fichier incorrect.

**Solution**: Utiliser des chemins Windows complets:
```bash
# ✅ Correct
claude "Extract VBA from C:\Users\alexi\Documents\budget.xlsm"

# ❌ Incorrect (chemin WSL)
claude "Extract VBA from /mnt/c/Users/alexi/Documents/budget.xlsm"
```

### Erreur: Timeout

**Cause**: Serveur surchargé ou fichier trop gros.

**Solution**:
1. Vérifier les logs serveur Windows
2. Réduire la taille du fichier
3. Redémarrer le serveur

### Firewall Windows

Si le client ne peut pas se connecter:

1. Ouvrir "Windows Defender Firewall"
2. "Advanced settings"
3. "Inbound Rules" → "New Rule"
4. Port: 8000, TCP
5. Allow the connection
6. Appliquer

## Configuration avancée

### Plusieurs clients

Plusieurs machines peuvent se connecter au même serveur:

**Windows Server** (192.168.1.100):
```powershell
vba-mcp-server-pro-http --host 0.0.0.0 --port 8000
```

**Client 1** (WSL):
```json
{
  "mcpServers": {
    "vba-pro": {"url": "http://192.168.1.100:8000/sse"}
  }
}
```

**Client 2** (macOS):
```json
{
  "mcpServers": {
    "vba-pro": {"url": "http://192.168.1.100:8000/sse"}
  }
}
```

### Port personnalisé

```powershell
# Windows
vba-mcp-server-pro-http --host 0.0.0.0 --port 9999
```

```json
// Client
{
  "mcpServers": {
    "vba-pro": {"url": "http://WIN_IP:9999/sse"}
  }
}
```

### Logs détaillés

```powershell
vba-mcp-server-pro-http --host 0.0.0.0 --port 8000 --log-level DEBUG
```

## Sécurité

### Réseau local (LAN)
- ✅ Sûr si vous faites confiance à votre réseau
- ⚠️ Pas d'authentification par défaut
- ⚠️ Accessible à tous sur le réseau local

### Recommandations
1. Utiliser sur réseau privé uniquement
2. Firewall pour restreindre les IPs
3. VPN pour accès distant
4. Ne PAS exposer sur Internet

## Scripts utiles

### Démarrage automatique (Windows Task Scheduler)

1. Ouvrir Task Scheduler
2. Create Basic Task
3. Trigger: At log on
4. Action: Start a program
5. Program: `C:\...\venv\Scripts\vba-mcp-server-pro-http.exe`
6. Arguments: `--host 0.0.0.0 --port 8000`

### Script de monitoring (PowerShell)

```powershell
# monitor_server.ps1
while ($true) {
    try {
        $response = Invoke-WebRequest "http://localhost:8000/sse" -TimeoutSec 2
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] ✓ Server OK" -ForegroundColor Green
    } catch {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] ✗ Server DOWN" -ForegroundColor Red
    }
    Start-Sleep -Seconds 30
}
```

## Exemples d'utilisation

### Extraction VBA

```bash
claude "Extract all VBA code from C:\Projects\budget.xlsm"
```

### Injection VBA

```bash
claude "Add this macro to C:\Projects\budget.xlsm:
Sub HelloWorld()
    MsgBox \"Hello from Claude!\"
End Sub"
```

### Analyse

```bash
claude "Analyze the code structure of C:\Projects\budget.xlsm"
```

### Tables Excel

```bash
claude "List all tables in C:\Data\sales.xlsx"
claude "Insert 10 rows at position 5 in the Sales table"
```

## Support

Pour des problèmes:
1. Vérifier les logs serveur (Windows)
2. Tester depuis Windows d'abord
3. Vérifier firewall et réseau
4. Email: alexistrouve.pro@gmail.com

## Comparaison: stdio vs HTTP/SSE

| Feature | stdio | HTTP/SSE |
|---------|-------|----------|
| **Setup** | Simple | Moyen |
| **OS Client** | Windows uniquement | Tous |
| **Performance** | Rapide | Légèrement plus lent |
| **Cas d'usage** | Local Windows | WSL, Mac, remote |
| **Sécurité** | Process isolation | Réseau (à sécuriser) |

## Next steps

1. ✅ Tester avec le script PowerShell
2. ✅ Configurer Claude Code
3. ✅ Tester une commande simple
4. ✅ Profiter de VBA automation cross-platform! 🎉
