"""
Office Session Manager

Manages persistent COM application sessions for interactive Office automation.
Keeps Office applications (Excel, Word, Access) open between MCP calls for faster operations.
"""

import asyncio
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Optional
import platform


# Configure logging
logger = logging.getLogger(__name__)


class OfficeSession:
    """Represents an open Office file session with COM application instance."""

    def __init__(
        self,
        app: Any,
        file_obj: Any,
        file_path: Path,
        app_type: str,
        read_only: bool = False
    ):
        """
        Initialize an Office session.

        Args:
            app: COM application instance (Excel.Application, Word.Application, etc.)
            file_obj: Document/Workbook/Database object
            file_path: Absolute path to the Office file
            app_type: Type of application ("Excel", "Word", "Access")
            read_only: Whether file is opened in read-only mode
        """
        self.app = app
        self.file_obj = file_obj
        self.file_path = file_path
        self.app_type = app_type
        self.read_only = read_only
        self.opened_at = datetime.now()
        self.last_accessed = datetime.now()
        self._vb_project = None  # Cached VBProject

    def is_alive(self) -> bool:
        """
        Check if COM objects are still valid.

        Returns:
            True if session is alive, False if COM objects are dead
        """
        try:
            # Test if app object still responds
            _ = self.app.Name
            if self.app_type == "Excel":
                _ = self.file_obj.Name
            elif self.app_type == "Word":
                _ = self.file_obj.Name
            elif self.app_type == "Access":
                _ = self.app.CurrentDb().Name
            return True
        except Exception:
            return False

    def refresh_last_accessed(self):
        """Update the last accessed timestamp."""
        self.last_accessed = datetime.now()

    @property
    def vb_project(self):
        """
        Get VBProject object (cached).

        Returns:
            VBProject COM object
        """
        if self._vb_project is None:
            try:
                if self.app_type in ["Excel", "Word"]:
                    self._vb_project = self.file_obj.VBProject
                elif self.app_type == "Access":
                    self._vb_project = self.app.VBE.ActiveVBProject
            except Exception as e:
                logger.warning(f"Failed to access VBProject: {e}")
        return self._vb_project

    def __repr__(self) -> str:
        return (
            f"<OfficeSession {self.app_type} "
            f"file={self.file_path.name} "
            f"read_only={self.read_only}>"
        )


class OfficeSessionManager:
    """
    Singleton manager for persistent Office COM application sessions.

    Maintains a pool of open Office files, automatically cleaning up stale sessions.
    """

    _instance: Optional["OfficeSessionManager"] = None
    _sessions: Dict[str, OfficeSession] = {}
    _lock: Optional[asyncio.Lock] = None
    _cleanup_task: Optional[asyncio.Task] = None

    # Configuration
    SESSION_TIMEOUT: int = 3600  # 1 hour in seconds
    CLEANUP_INTERVAL: int = 300  # 5 minutes in seconds

    def __new__(cls):
        """Singleton pattern."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        """Initialize the session manager (only once due to singleton)."""
        if self._initialized:
            return
        self._sessions = {}
        self._lock = asyncio.Lock()
        self._cleanup_task = None
        self._initialized = True
        logger.info("OfficeSessionManager initialized")

    @classmethod
    def get_instance(cls) -> "OfficeSessionManager":
        """
        Get the singleton instance.

        Returns:
            OfficeSessionManager instance
        """
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    async def get_or_create_session(
        self,
        file_path: Path,
        read_only: bool = False,
        force_new: bool = False
    ) -> OfficeSession:
        """
        Get existing session or create a new one.

        Args:
            file_path: Absolute path to Office file
            read_only: Open in read-only mode
            force_new: Force creation of new session (close existing if present)

        Returns:
            OfficeSession instance

        Raises:
            FileNotFoundError: If file doesn't exist
            RuntimeError: If not on Windows
            PermissionError: If file is locked
        """
        # Normalize path
        file_path = file_path.resolve()
        file_key = str(file_path)

        async with self._lock:
            # Check for existing session
            if file_key in self._sessions and not force_new:
                session = self._sessions[file_key]

                # Check if session is still alive
                if session.is_alive():
                    logger.info(f"Reusing existing session for {file_path.name}")
                    session.refresh_last_accessed()
                    return session
                else:
                    logger.warning(f"Stale session detected for {file_path.name}, recreating")
                    await self._force_close_session(file_key)

            # Check if file is locked before opening
            if self._check_file_lock(file_path):
                # If file is locked, check if it's in our sessions
                if file_key in self._sessions:
                    session = self._sessions[file_key]
                    if session.is_alive():
                        # Our session is alive, reuse it
                        session.refresh_last_accessed()
                        return session
                    else:
                        # Session is dead but file is locked - external process
                        raise PermissionError(
                            f"File is locked by another process: {file_path}\n"
                            f"The file appears to be open in another application."
                        )
                else:
                    # Not in our sessions - locked by external process
                    raise PermissionError(
                        f"File is locked by another application: {file_path.name}\n"
                        f"Close the file in Excel/Word and try again."
                    )

            # Create new session
            logger.info(f"Creating new session for {file_path.name}")
            session = await self._create_session(file_path, read_only)
            self._sessions[file_key] = session
            return session

    async def _create_session(self, file_path: Path, read_only: bool) -> OfficeSession:
        """
        Create a new Office session.

        Args:
            file_path: Absolute path to Office file
            read_only: Open in read-only mode

        Returns:
            New OfficeSession instance

        Raises:
            FileNotFoundError: If file doesn't exist
            RuntimeError: If not on Windows or missing dependencies
            PermissionError: If file is locked
        """
        # Validate file exists
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        # Platform check
        if platform.system() != "Windows":
            raise RuntimeError(
                "Office automation is only supported on Windows. "
                "Install this package on a Windows machine with Microsoft Office."
            )

        # Import Windows-specific modules
        try:
            import win32com.client
            import pythoncom
        except ImportError:
            raise RuntimeError(
                "pywin32 is required for Office automation. "
                "Install with: pip install vba-mcp-server-pro[windows]"
            )

        # Detect application type
        file_ext = file_path.suffix.lower()
        if file_ext in ['.xlsm', '.xlsb', '.xls', '.xlsx']:
            app_type = "Excel"
            app_name = "Excel.Application"
        elif file_ext in ['.docm', '.doc', '.docx']:
            app_type = "Word"
            app_name = "Word.Application"
        elif file_ext in ['.accdb', '.mdb']:
            app_type = "Access"
            app_name = "Access.Application"
        else:
            raise ValueError(f"Unsupported file type: {file_ext}")

        # Initialize COM
        pythoncom.CoInitialize()

        try:
            # Create COM application
            app = win32com.client.Dispatch(app_name)

            # Try to set Visible=True, but continue if it fails (WSL compatibility)
            try:
                app.Visible = True  # Always visible for interactive use
                logger.debug("App.Visible set to True (interactive mode)")
            except Exception as e:
                logger.warning(f"Could not set App.Visible=True: {e}. Continuing anyway.")

            # Try to set DisplayAlerts=False, but continue if it fails
            try:
                app.DisplayAlerts = False  # Suppress prompts
                logger.debug("App.DisplayAlerts set to False")
            except Exception as e:
                logger.warning(f"Could not set App.DisplayAlerts=False: {e}")

            # Open file
            file_obj = self._open_file(app, app_type, file_path, read_only)

            # Create session object
            session = OfficeSession(
                app=app,
                file_obj=file_obj,
                file_path=file_path,
                app_type=app_type,
                read_only=read_only
            )

            logger.info(f"Session created successfully for {file_path.name} ({app_type})")
            return session

        except Exception as e:
            # Cleanup on error
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            raise

    def _open_file(
        self,
        app: Any,
        app_type: str,
        file_path: Path,
        read_only: bool
    ) -> Any:
        """
        Open Office file with application-specific method.

        Args:
            app: COM application instance
            app_type: "Excel", "Word", or "Access"
            file_path: Path to file
            read_only: Read-only mode

        Returns:
            File object (Workbook, Document, etc.)

        Raises:
            PermissionError: If file is locked
        """
        try:
            import pythoncom

            if app_type == "Excel":
                file_obj = app.Workbooks.Open(
                    str(file_path),
                    ReadOnly=read_only,
                    UpdateLinks=False
                )
            elif app_type == "Word":
                file_obj = app.Documents.Open(
                    str(file_path),
                    ReadOnly=read_only
                )
            elif app_type == "Access":
                # Access has different method
                exclusive = not read_only
                app.OpenCurrentDatabase(str(file_path), Exclusive=exclusive)
                file_obj = app  # Access uses app itself as file object
            else:
                raise ValueError(f"Unknown app type: {app_type}")

            return file_obj

        except pythoncom.com_error as e:
            error_msg = str(e).lower()
            if "in use" in error_msg or "locked" in error_msg or "opened by another user" in error_msg:
                raise PermissionError(
                    f"File is already open in another application: {file_path.name}\n"
                    f"Close the file in {app_type} and try again."
                )
            raise RuntimeError(f"Failed to open file: {str(e)}")

    async def close_session(self, file_path: Path, save: bool = True) -> None:
        """
        Close an Office session.

        Args:
            file_path: Absolute path to Office file
            save: Save changes before closing

        Raises:
            ValueError: If session doesn't exist
        """
        file_path = file_path.resolve()
        file_key = str(file_path)

        async with self._lock:
            if file_key not in self._sessions:
                raise ValueError(f"No open session for {file_path.name}")

            session = self._sessions[file_key]
            await self._close_session_internal(session, save)
            del self._sessions[file_key]
            logger.info(f"Session closed for {file_path.name}")

    async def _close_session_internal(self, session: OfficeSession, save: bool = True) -> None:
        """
        Internal method to close a session.

        Args:
            session: OfficeSession to close
            save: Save changes before closing
        """
        try:
            import pythoncom

            # Save if requested
            if save and not session.read_only:
                try:
                    if session.app_type in ["Excel", "Word"]:
                        session.file_obj.Save()
                    elif session.app_type == "Access":
                        # Access auto-saves
                        pass
                except Exception as e:
                    logger.warning(f"Failed to save file: {e}")

            # Close file
            try:
                if session.app_type == "Excel":
                    session.file_obj.Close(SaveChanges=False)  # Already saved above
                elif session.app_type == "Word":
                    session.file_obj.Close(SaveChanges=False)
                elif session.app_type == "Access":
                    session.app.CloseCurrentDatabase()
            except Exception as e:
                logger.warning(f"Failed to close file: {e}")

            # Quit application
            try:
                session.app.Quit()
            except Exception as e:
                logger.warning(f"Failed to quit application: {e}")

        finally:
            # Release COM objects before CoUninitialize
            self._release_com_objects(session)

            # Always uninitialize COM
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception as e:
                logger.warning(f"Failed to uninitialize COM: {e}")

    async def _force_close_session(self, file_key: str) -> None:
        """
        Force close a session without saving (for cleanup).

        Args:
            file_key: Session key (normalized path string)
        """
        if file_key in self._sessions:
            session = self._sessions[file_key]
            await self._close_session_internal(session, save=False)
            del self._sessions[file_key]

    async def close_all_sessions(self, save: bool = True) -> None:
        """
        Close all open sessions.

        Args:
            save: Save changes before closing
        """
        logger.info(f"Closing all sessions (save={save})")
        async with self._lock:
            session_keys = list(self._sessions.keys())
            for file_key in session_keys:
                session = self._sessions[file_key]
                await self._close_session_internal(session, save)
                del self._sessions[file_key]
        logger.info("All sessions closed")

    async def _cleanup_stale_sessions(self) -> None:
        """Background task to cleanup stale sessions."""
        logger.info("Cleanup task started")
        try:
            while True:
                await asyncio.sleep(self.CLEANUP_INTERVAL)

                async with self._lock:
                    now = datetime.now()
                    stale_keys = []

                    for file_key, session in self._sessions.items():
                        # Check if session is stale (timeout exceeded)
                        age = now - session.last_accessed
                        if age.total_seconds() > self.SESSION_TIMEOUT:
                            logger.info(
                                f"Session for {session.file_path.name} is stale "
                                f"({age.total_seconds():.0f}s old), closing"
                            )
                            stale_keys.append(file_key)
                        # Also check if session is dead
                        elif not session.is_alive():
                            logger.warning(
                                f"Session for {session.file_path.name} is dead, removing"
                            )
                            stale_keys.append(file_key)

                    # Close stale sessions
                    for file_key in stale_keys:
                        session = self._sessions[file_key]
                        await self._close_session_internal(session, save=True)
                        del self._sessions[file_key]

                    if stale_keys:
                        logger.info(f"Cleaned up {len(stale_keys)} stale session(s)")

        except asyncio.CancelledError:
            logger.info("Cleanup task cancelled")
            raise
        except Exception as e:
            logger.error(f"Cleanup task error: {e}", exc_info=True)

    def start_cleanup_task(self) -> None:
        """Start the background cleanup task."""
        if self._cleanup_task is None or self._cleanup_task.done():
            self._cleanup_task = asyncio.create_task(self._cleanup_stale_sessions())
            logger.info("Cleanup task started")
        else:
            logger.warning("Cleanup task already running")

    async def stop_cleanup_task(self) -> None:
        """Stop the background cleanup task."""
        if self._cleanup_task is not None and not self._cleanup_task.done():
            self._cleanup_task.cancel()
            try:
                await self._cleanup_task
            except asyncio.CancelledError:
                pass
            logger.info("Cleanup task stopped")

    def _release_com_objects(self, session: OfficeSession) -> None:
        """
        Release COM objects explicitly to prevent file locks and memory leaks.

        Args:
            session: OfficeSession to release COM objects from
        """
        try:
            import pythoncom

            # Release VBProject if it exists
            if session._vb_project is not None:
                try:
                    pythoncom.ReleaseObject(session._vb_project)
                    session._vb_project = None
                except Exception as e:
                    logger.warning(f"Error releasing VBProject: {e}")

            # Release file object
            if session.file_obj is not None:
                try:
                    pythoncom.ReleaseObject(session.file_obj)
                except Exception as e:
                    logger.warning(f"Error releasing file object: {e}")

            # Release application object
            if session.app is not None:
                try:
                    pythoncom.ReleaseObject(session.app)
                except Exception as e:
                    logger.warning(f"Error releasing app object: {e}")

        except Exception as e:
            logger.warning(f"Error during COM cleanup: {e}")

    def _check_file_lock(self, file_path: Path) -> bool:
        """
        Check if file is locked by another process.

        Args:
            file_path: Path to file to check

        Returns:
            True if file is locked, False if accessible
        """
        try:
            import win32file
            import pywintypes

            try:
                # Try to open file with exclusive access
                handle = win32file.CreateFile(
                    str(file_path),
                    win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                    0,  # No sharing - exclusive access
                    None,
                    win32file.OPEN_EXISTING,
                    0,
                    None
                )
                # If successful, file is not locked - close handle
                win32file.CloseHandle(handle)
                return False
            except pywintypes.error:
                # File is locked
                return True

        except ImportError:
            # If win32file not available, assume not locked
            logger.warning("win32file not available, cannot check file lock")
            return False
        except Exception as e:
            # On error, assume not locked to avoid blocking valid operations
            logger.warning(f"Error checking file lock: {e}")
            return False

    def list_sessions(self) -> list[Dict[str, Any]]:
        """
        List all active sessions.

        Returns:
            List of session info dicts
        """
        sessions_info = []
        for file_key, session in self._sessions.items():
            age = datetime.now() - session.opened_at
            sessions_info.append({
                "file_name": session.file_path.name,
                "file_path": str(session.file_path),
                "app_type": session.app_type,
                "read_only": session.read_only,
                "age_seconds": age.total_seconds(),
                "last_accessed_seconds": (datetime.now() - session.last_accessed).total_seconds()
            })
        return sessions_info
