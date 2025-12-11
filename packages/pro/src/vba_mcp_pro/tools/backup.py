"""
VBA Backup Tool (PRO)

Manages backups of Office files before modifications.
"""

from pathlib import Path
from typing import List, Optional
from datetime import datetime
import shutil
import json


BACKUP_MANIFEST = ".vba_backups.json"


async def backup_tool(
    file_path: str,
    action: str = "create",
    backup_id: Optional[str] = None
) -> str:
    """
    Manage backups of Office files.

    Args:
        file_path: Absolute path to Office file
        action: Action to perform (create, list, restore, delete)
        backup_id: Backup identifier for restore/delete actions

    Returns:
        Action result message

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If action is invalid or backup not found
    """
    path = Path(file_path)

    if action == "create":
        return await _create_backup(path)
    elif action == "list":
        return await _list_backups(path)
    elif action == "restore":
        if not backup_id:
            raise ValueError("backup_id required for restore action")
        return await _restore_backup(path, backup_id)
    elif action == "delete":
        if not backup_id:
            raise ValueError("backup_id required for delete action")
        return await _delete_backup(path, backup_id)
    else:
        raise ValueError(f"Unknown action: {action}. Use: create, list, restore, delete")


async def _create_backup(file_path: Path) -> str:
    """Create a new backup."""
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Generate backup info
    timestamp = datetime.now()
    backup_id = timestamp.strftime("%Y%m%d_%H%M%S")
    backup_name = f"{file_path.stem}_backup_{backup_id}{file_path.suffix}"
    backup_dir = file_path.parent / ".vba_backups"
    backup_dir.mkdir(exist_ok=True)
    backup_path = backup_dir / backup_name

    # Copy file
    shutil.copy2(file_path, backup_path)

    # Update manifest
    manifest = _load_manifest(file_path)
    manifest["backups"].append({
        "id": backup_id,
        "filename": backup_name,
        "created": timestamp.isoformat(),
        "original_size": file_path.stat().st_size
    })
    _save_manifest(file_path, manifest)

    return "\n".join([
        f"**Backup Created**",
        f"",
        f"File: {file_path.name}",
        f"Backup ID: {backup_id}",
        f"Location: {backup_path}",
        f"",
        f"To restore: Use action='restore' with backup_id='{backup_id}'"
    ])


async def _list_backups(file_path: Path) -> str:
    """List all backups for a file."""
    manifest = _load_manifest(file_path)
    backups = manifest.get("backups", [])

    if not backups:
        return f"No backups found for {file_path.name}"

    lines = [
        f"**Backups for {file_path.name}**",
        "",
        f"Total: {len(backups)} backup(s)",
        ""
    ]

    for b in reversed(backups):  # Most recent first
        lines.append(f"- **{b['id']}** - {b['created'][:16]} ({b['original_size']:,} bytes)")

    return "\n".join(lines)


async def _restore_backup(file_path: Path, backup_id: str) -> str:
    """Restore from a backup."""
    manifest = _load_manifest(file_path)
    backup = next((b for b in manifest["backups"] if b["id"] == backup_id), None)

    if not backup:
        raise ValueError(f"Backup '{backup_id}' not found")

    backup_dir = file_path.parent / ".vba_backups"
    backup_path = backup_dir / backup["filename"]

    if not backup_path.exists():
        raise FileNotFoundError(f"Backup file missing: {backup_path}")

    # Create safety backup of current file before restore
    if file_path.exists():
        safety_name = f"{file_path.stem}_pre_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}{file_path.suffix}"
        shutil.copy2(file_path, backup_dir / safety_name)

    # Restore
    shutil.copy2(backup_path, file_path)

    return "\n".join([
        f"**Backup Restored**",
        f"",
        f"File: {file_path.name}",
        f"Restored from: {backup_id}",
        f"Backup date: {backup['created'][:16]}"
    ])


async def _delete_backup(file_path: Path, backup_id: str) -> str:
    """Delete a backup."""
    manifest = _load_manifest(file_path)
    backup = next((b for b in manifest["backups"] if b["id"] == backup_id), None)

    if not backup:
        raise ValueError(f"Backup '{backup_id}' not found")

    backup_dir = file_path.parent / ".vba_backups"
    backup_path = backup_dir / backup["filename"]

    # Delete file
    if backup_path.exists():
        backup_path.unlink()

    # Update manifest
    manifest["backups"] = [b for b in manifest["backups"] if b["id"] != backup_id]
    _save_manifest(file_path, manifest)

    return f"Backup '{backup_id}' deleted successfully"


def _load_manifest(file_path: Path) -> dict:
    """Load or create backup manifest."""
    backup_dir = file_path.parent / ".vba_backups"
    manifest_path = backup_dir / BACKUP_MANIFEST

    if manifest_path.exists():
        with open(manifest_path, "r") as f:
            return json.load(f)

    return {
        "file": str(file_path.name),
        "backups": []
    }


def _save_manifest(file_path: Path, manifest: dict):
    """Save backup manifest."""
    backup_dir = file_path.parent / ".vba_backups"
    backup_dir.mkdir(exist_ok=True)
    manifest_path = backup_dir / BACKUP_MANIFEST

    with open(manifest_path, "w") as f:
        json.dump(manifest, f, indent=2)
