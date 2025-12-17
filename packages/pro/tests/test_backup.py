"""Tests for backup tool."""
import pytest
from pathlib import Path
import shutil

from vba_mcp_pro.tools.backup import (
    backup_tool,
    _create_backup,
    _list_backups,
    _restore_backup,
    _delete_backup,
    _load_manifest,
    _save_manifest
)


class TestBackupManagement:
    """Test suite for backup management."""

    @pytest.mark.asyncio
    async def test_create_backup(self, sample_xlsm_copy):
        """Test creating a backup."""
        result = await _create_backup(sample_xlsm_copy)

        assert "Backup Created" in result
        assert "Backup ID:" in result

        # Check backup directory exists
        backup_dir = sample_xlsm_copy.parent / ".vba_backups"
        assert backup_dir.exists()

        # Check manifest exists
        manifest = _load_manifest(sample_xlsm_copy)
        assert len(manifest["backups"]) == 1

    @pytest.mark.asyncio
    async def test_create_backup_file_not_found(self, tmp_path):
        """Test creating backup for non-existent file."""
        fake_file = tmp_path / "nonexistent.xlsm"
        with pytest.raises(FileNotFoundError):
            await _create_backup(fake_file)

    @pytest.mark.asyncio
    async def test_list_backups_empty(self, sample_xlsm):
        """Test listing backups when none exist."""
        result = await _list_backups(sample_xlsm)
        assert "No backups found" in result

    @pytest.mark.asyncio
    async def test_list_backups_with_backups(self, sample_xlsm_copy):
        """Test listing existing backups."""
        # Create some backups
        await _create_backup(sample_xlsm_copy)
        await _create_backup(sample_xlsm_copy)

        result = await _list_backups(sample_xlsm_copy)

        assert "Backups for" in result
        assert "Total: 2 backup(s)" in result

    @pytest.mark.asyncio
    async def test_restore_backup(self, sample_xlsm_copy):
        """Test restoring from backup."""
        # Create initial backup
        await _create_backup(sample_xlsm_copy)

        # Get backup ID
        manifest = _load_manifest(sample_xlsm_copy)
        backup_id = manifest["backups"][0]["id"]

        # Modify the file (simulate changes)
        original_size = sample_xlsm_copy.stat().st_size
        with open(sample_xlsm_copy, "ab") as f:
            f.write(b"MODIFIED")

        # Restore
        result = await _restore_backup(sample_xlsm_copy, backup_id)

        assert "Backup Restored" in result
        assert backup_id in result

        # Check file was restored
        restored_size = sample_xlsm_copy.stat().st_size
        assert restored_size == original_size

    @pytest.mark.asyncio
    async def test_restore_backup_not_found(self, sample_xlsm_copy):
        """Test restoring non-existent backup."""
        with pytest.raises(ValueError, match="not found"):
            await _restore_backup(sample_xlsm_copy, "nonexistent_id")

    @pytest.mark.asyncio
    async def test_delete_backup(self, sample_xlsm_copy):
        """Test deleting a backup."""
        # Create backup
        await _create_backup(sample_xlsm_copy)

        # Get backup ID
        manifest = _load_manifest(sample_xlsm_copy)
        backup_id = manifest["backups"][0]["id"]

        # Delete backup
        result = await _delete_backup(sample_xlsm_copy, backup_id)

        assert "deleted successfully" in result

        # Verify deleted from manifest
        manifest = _load_manifest(sample_xlsm_copy)
        assert len(manifest["backups"]) == 0

    @pytest.mark.asyncio
    async def test_backup_tool_actions(self, sample_xlsm_copy):
        """Test backup_tool with different actions."""
        # Create
        result = await backup_tool(str(sample_xlsm_copy), action="create")
        assert "Backup Created" in result

        # List
        result = await backup_tool(str(sample_xlsm_copy), action="list")
        assert "Backups for" in result

        # Get backup ID
        manifest = _load_manifest(sample_xlsm_copy)
        backup_id = manifest["backups"][0]["id"]

        # Restore
        result = await backup_tool(str(sample_xlsm_copy), action="restore", backup_id=backup_id)
        assert "Backup Restored" in result

        # Delete
        result = await backup_tool(str(sample_xlsm_copy), action="delete", backup_id=backup_id)
        assert "deleted successfully" in result

    @pytest.mark.asyncio
    async def test_backup_tool_invalid_action(self, sample_xlsm):
        """Test backup_tool with invalid action."""
        with pytest.raises(ValueError, match="Unknown action"):
            await backup_tool(str(sample_xlsm), action="invalid")

    def test_manifest_operations(self, tmp_path):
        """Test manifest load/save operations."""
        test_file = tmp_path / "test.xlsm"
        test_file.touch()

        # Load non-existent manifest
        manifest = _load_manifest(test_file)
        assert manifest["backups"] == []

        # Add backup entry
        manifest["backups"].append({
            "id": "test_id",
            "filename": "test.xlsm",
            "created": "2024-01-01T00:00:00",
            "original_size": 1000
        })

        # Save manifest
        _save_manifest(test_file, manifest)

        # Load again
        loaded = _load_manifest(test_file)
        assert len(loaded["backups"]) == 1
        assert loaded["backups"][0]["id"] == "test_id"
