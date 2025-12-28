# VBA MCP Server Pro - HTTP/SSE Setup Guide

## Overview

The HTTP/SSE transport allows you to run the VBA MCP server on Windows (where COM automation works) and connect to it from any OS (WSL, macOS, Linux).

## Architecture

```
┌─────────────────┐                    ┌──────────────────┐
│   WSL/Mac/Linux │                    │   Windows        │
│   (Client)      │ ◄─── HTTP/SSE ───► │   (Server)       │
│                 │                    │                  │
│  Claude Code    │                    │  VBA MCP Pro     │
│  Claude Code    │                    │  + COM           │
└─────────────────┘                    └──────────────────┘
```

## Quick Start

### 1. Install Dependencies

On **Windows** (where the server will run):

```bash
# Navigate to the monorepo
cd /mnt/c/Users/alexi/Documents/projects/vba-mcp-monorepo

# Activate virtual environment
source venv/bin/activate  # or venv\Scripts\activate on Windows

# Reinstall Pro package with new dependencies
pip install -e packages/pro[windows]
```

### 2. Start the Server

On **Windows**:

```bash
# Start server on localhost (accessible only from same machine)
vba-mcp-server-pro-http

# OR start on all interfaces (accessible from network)
vba-mcp-server-pro-http --host 0.0.0.0 --port 8000
```

The server will start on:
- SSE endpoint: `http://localhost:8000/sse`
- Messages endpoint: `http://localhost:8000/messages/`

### 3. Configure Client

#### Claude Code (macOS/Windows/WSL)

Edit your Claude Code config:

**macOS**: `~/.claude/config.json`
**Windows**: `%USERPROFILE%\.claude\config.json`
**WSL**: `~/.claude/config.json`

```json
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://localhost:8000/sse"
    }
  }
}
```

#### Claude Code (CLI)

Create/edit `~/.claude/config.json`:

```json
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://localhost:8000/sse"
    }
  }
}
```

### 4. Test Connection

Restart Claude Code, then try:

```
"List all VBA modules in C:\Users\alexi\Documents\test.xlsm"
```

## Network Access

### Same Machine (localhost)

If client and server are on the same Windows machine:

```bash
vba-mcp-server-pro-http --host 127.0.0.1 --port 8000
```

Client config:
```json
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://localhost:8000/sse"
    }
  }
}
```

### WSL to Windows

If client is in WSL and server is on Windows host:

**On Windows**:
```bash
# Get Windows IP
ipconfig

# Start server on all interfaces
vba-mcp-server-pro-http --host 0.0.0.0 --port 8000
```

**In WSL**:

Find Windows host IP:
```bash
# Method 1: From /etc/resolv.conf
cat /etc/resolv.conf | grep nameserver | awk '{print $2}'

# Method 2: Usually 172.x.x.1
# Example: 172.28.160.1
```

Client config (WSL):
```json
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://172.28.160.1:8000/sse"
    }
  }
}
```

### Network Access (Different Machines)

**On Windows Server**:
```bash
# Get your IP address
ipconfig

# Example output: 192.168.1.100

# Start server
vba-mcp-server-pro-http --host 0.0.0.0 --port 8000
```

**On Client Machine**:
```json
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://192.168.1.100:8000/sse"
    }
  }
}
```

## Command Line Options

```bash
vba-mcp-server-pro-http --help

Options:
  --host HOST          Host to bind to (default: 127.0.0.1)
                       Use 0.0.0.0 for all interfaces
  --port PORT          Port to listen on (default: 8000)
  --log-level LEVEL    Logging level (DEBUG, INFO, WARNING, ERROR)
```

## Security Considerations

### localhost (127.0.0.1)
- ✅ Safe - Only accessible from same machine
- Use for local development and testing

### All interfaces (0.0.0.0)
- ⚠️ Caution - Accessible from network
- Only use on trusted networks
- Consider firewall rules
- No built-in authentication (trust network)

### Production
For production use, consider:
- Reverse proxy with authentication (nginx, Apache)
- VPN for remote access
- Firewall rules to restrict access
- HTTPS with TLS certificates

## Troubleshooting

### Connection Refused

**Problem**: Client can't connect to server

**Solutions**:
1. Check server is running: `netstat -an | grep 8000`
2. Check firewall allows port 8000
3. Verify IP address is correct
4. Try `--host 0.0.0.0` on server

### Server Crashes

**Problem**: Server exits with error

**Solutions**:
1. Check logs: `vba-mcp-server-pro-http --log-level DEBUG`
2. Verify pywin32 installed: `pip list | grep pywin32`
3. Run on native Windows (not WSL) for COM

### Tools Not Working

**Problem**: Client connects but tools fail

**Solutions**:
1. Verify Windows paths (use `C:\` not `/mnt/c/`)
2. Check Excel/Office installed on server machine
3. Verify file permissions

## Comparison: stdio vs HTTP/SSE

| Feature | stdio (Default) | HTTP/SSE |
|---------|----------------|----------|
| **Client Location** | Same machine | Any network location |
| **OS Support** | Client must be Windows | Client can be any OS |
| **Setup Complexity** | Simple | Medium |
| **Performance** | Fastest | Slightly slower (network) |
| **Security** | Process isolation | Network security needed |
| **Use Case** | Local development | Remote access, WSL, cross-OS |

## Examples

### Development Setup (WSL + Windows)

**On Windows** (PowerShell):
```powershell
cd C:\Users\alexi\Documents\projects\vba-mcp-monorepo
venv\Scripts\activate
vba-mcp-server-pro-http --host 0.0.0.0
```

**In WSL**:
```bash
# Get Windows IP
WIN_IP=$(cat /etc/resolv.conf | grep nameserver | awk '{print $2}')
echo "Windows IP: $WIN_IP"

# Configure Claude Code
cat > ~/.claude/config.json <<EOF
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://${WIN_IP}:8000/sse"
    }
  }
}
EOF

# Test
claude-code "List modules in /mnt/c/Users/alexi/Documents/test.xlsm"
```

### Team Setup (Shared Server)

One Windows machine runs the server, team members connect:

**Server** (192.168.1.100):
```bash
vba-mcp-server-pro-http --host 0.0.0.0 --port 8000
```

**Team Member 1** (Windows):
```json
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://192.168.1.100:8000/sse"
    }
  }
}
```

**Team Member 2** (macOS):
```json
{
  "mcpServers": {
    "vba-pro": {
      "url": "http://192.168.1.100:8000/sse"
    }
  }
}
```

## Monitoring

### Check Server Status

```bash
# View logs
vba-mcp-server-pro-http --log-level INFO

# Test endpoint
curl http://localhost:8000/sse
```

### Monitor Connections

Server logs will show:
- Client connections
- Tool calls
- Errors
- Session management

## Advanced: Systemd Service (Linux/WSL)

Create `/etc/systemd/system/vba-mcp-pro.service`:

```ini
[Unit]
Description=VBA MCP Server Pro (HTTP/SSE)
After=network.target

[Service]
Type=simple
User=your-user
WorkingDirectory=/path/to/vba-mcp-monorepo
ExecStart=/path/to/venv/bin/vba-mcp-server-pro-http --host 0.0.0.0 --port 8000
Restart=always

[Install]
WantedBy=multi-user.target
```

Enable and start:
```bash
sudo systemctl enable vba-mcp-pro
sudo systemctl start vba-mcp-pro
sudo systemctl status vba-mcp-pro
```

## Next Steps

1. Start the server: `vba-mcp-server-pro-http`
2. Configure your client with the server URL
3. Test with a simple query
4. Enjoy cross-platform VBA automation!

## Support

For issues or questions:
- Email: alexistrouve.pro@gmail.com
- Check server logs with `--log-level DEBUG`
