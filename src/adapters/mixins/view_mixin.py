"""
View mixin for AutoCAD adapter.

Handles view operations (zoom, refresh, undo, redo).
"""

import logging
import base64
import os
import time
import re
import win32gui
import win32con
from typing import TYPE_CHECKING, Dict
from PIL import ImageGrab

logger = logging.getLogger(__name__)


class ViewMixin:
    """Mixin for view operations."""

    if TYPE_CHECKING:
        # Tell type checker this mixin is used with CADAdapterProtocol
        from typing import Any

        # Attributes from AutoCADAdapter
        cad_type: str

        def _get_application(self, operation: str = "operation") -> Any: ...
        def _get_document(self, operation: str = "operation") -> Any: ...
        def _simulate_autocad_click(self) -> bool: ...
        def _validate_connection(self) -> None: ...
        def resolve_export_path(
            self, filename: str, folder_type: str = "drawings"
        ) -> str: ...

    def _sanitize_command_input(self, user_input: str) -> str:
        """Sanitize input for SendCommand to prevent command injection.

        Restricts input to safe characters that are common in file paths and CAD commands.
        Non-matching characters are removed.

        Args:
            user_input: The user-provided input to sanitize

        Returns:
            str: The sanitized input safe for SendCommand
        """
        safe_pattern = re.compile(r"^[a-zA-Z0-9\s\\/._\-:()]+$")
        if not safe_pattern.match(user_input):
            logger.warning(f"Input sanitized due to unsafe characters: {user_input}")
            # Remove all characters that don't match safe pattern
            sanitized = re.sub(r"[^a-zA-Z0-9\s\\/._\-:()]", "", user_input)
            logger.debug(f"Sanitized to: {sanitized}")
            return sanitized
        return user_input

    def _find_cad_window(self) -> int:
        """Find CAD application window.

        Collects all top-level windows whose title contains the CAD search
        term.  Prefers windows whose class name matches a known CAD frame
        class, but falls back to title-only matching when the class list
        does not cover the installed CAD version (e.g. AutoCAD 2020+ uses
        Afx:… class names).  Among all candidates the window with the
        largest area is chosen to avoid selecting dialogs or sub-windows.

        Returns:
            int: Window handle (HWND) of the main CAD window.

        Raises:
            Exception: If no matching window is found.
        """
        from mcp_tools.constants import CAD_WINDOW_SEARCH_TERMS, AUTOCAD_WINDOW_CLASSES

        search_term = CAD_WINDOW_SEARCH_TERMS.get(self.cad_type, "").lower()

        strict: list  = []  # title + known class
        loose:  list  = []  # title only (fallback)

        def enum_windows_callback(h, _):
            title = win32gui.GetWindowText(h)
            if not title or search_term not in title.lower():
                return
            if "VBA" in title:
                return

            class_name = win32gui.GetClassName(h)
            # For minimized windows GetWindowRect returns a tiny stub rect;
            # treat them as large so they are not excluded.
            if win32gui.IsIconic(h):
                area = 10 ** 9
            else:
                rect = win32gui.GetWindowRect(h)
                area = max(0, rect[2] - rect[0]) * max(0, rect[3] - rect[1])
            if area < 100:            # skip genuinely tiny dialogs / tooltips
                return

            if any(p in class_name for p in AUTOCAD_WINDOW_CLASSES):
                strict.append((area, h))
            else:
                loose.append((area, h))

            logger.debug(
                f"CAD candidate: title='{title}', class='{class_name}', "
                f"hwnd={h}, area={area}"
            )

        win32gui.EnumWindows(enum_windows_callback, None)

        candidates = strict or loose
        if not candidates:
            raise Exception(
                f"Could not find window for {self.cad_type} "
                f"(search term: '{search_term}')"
            )

        _, hwnd = max(candidates)   # pick largest window
        logger.debug(f"Selected CAD window hwnd={hwnd}")
        return hwnd

    def get_screenshot(self) -> Dict[str, str]:
        """
        Capture a screenshot of the CAD application window.

        Returns:
            dict: Dictionary with 'path' and 'data' (base64)

        Raises:
            Exception: If screenshot fails
        """
        try:
            self._validate_connection()

            hwnd = self._find_cad_window()

            import ctypes
            import win32ui
            import time

            # If the window is minimized, restore it without activating
            # (SW_SHOWNOACTIVATE = 4) so PrintWindow gets a full render.
            # We put it back to minimized afterwards.
            was_minimized = win32gui.IsIconic(hwnd)
            if was_minimized:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
                time.sleep(0.3)   # give the compositor time to render

            rect = win32gui.GetWindowRect(hwnd)
            w, h = rect[2] - rect[0], rect[3] - rect[1]
            logger.debug(f"Capturing via PrintWindow: HWND {hwnd}, {w}x{h}")

            import ctypes
            import win32ui

            hdc_win = win32gui.GetWindowDC(hwnd)
            dc_obj = win32ui.CreateDCFromHandle(hdc_win)
            mem_dc = dc_obj.CreateCompatibleDC()

            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(dc_obj, w, h)
            mem_dc.SelectObject(bmp)

            # PW_RENDERFULLCONTENT (2) captures composited/DPI-aware content
            result = ctypes.windll.user32.PrintWindow(hwnd, mem_dc.GetSafeHdc(), 2)
            if not result:
                # Fallback to legacy PrintWindow without flags
                ctypes.windll.user32.PrintWindow(hwnd, mem_dc.GetSafeHdc(), 0)

            bmp_info = bmp.GetInfo()
            bmp_bits = bmp.GetBitmapBits(True)

            from PIL import Image
            image = Image.frombuffer(
                "RGB",
                (bmp_info["bmWidth"], bmp_info["bmHeight"]),
                bmp_bits,
                "raw",
                "BGRX",
                0,
                1,
            )

            mem_dc.DeleteDC()
            dc_obj.DeleteDC()
            win32gui.ReleaseDC(hwnd, hdc_win)
            win32gui.DeleteObject(bmp.GetHandle())

            # Restore minimized state if we changed it
            if was_minimized:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)

            filename = f"cad_screenshot_{os.getpid()}.png"
            filepath = self.resolve_export_path(filename, "images")
            image.save(filepath, "PNG")

            with open(filepath, "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode("utf-8")

            logger.info(f"Screenshot saved to {filepath}")
            return {"path": filepath, "data": encoded_string}

        except Exception as e:
            logger.error(f"Screenshot failed: {e}")
            raise Exception(f"Failed to capture screenshot: {e}")

    def export_view(self) -> Dict[str, str]:
        """Export current view using internal PNGOUT command.

        This method uses ZWCAD's built-in PNGOUT command to export the drawing
        view to a PNG file. Unlike get_screenshot(), this method:
        - Works even if the window is minimized or obscured
        - Captures only the drawing content (no UI chrome)
        - Uses the CAD application's internal rendering

        Returns:
            Dictionary with 'path' and 'data' (base64 encoded image)

        Raises:
            Exception: If export fails
        """
        try:
            self._validate_connection()
            document = self._get_document("export_view")

            # Prepare filename and resolve path using centralized utility
            filename = f"cad_export_{os.getpid()}.png"
            filepath = self.resolve_export_path(filename, "images")

            # Ensure path uses backslashes for CAD command and is absolute
            filepath_cad = filepath.replace("/", "\\")

            logger.debug(f"Exporting view to {filepath_cad}")

            # Find the CAD window for focusing (with error handling for headless mode)
            try:
                hwnd = self._find_cad_window()
                try:
                    if win32gui.IsIconic(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.5)
                except Exception as e:
                    logger.warning(f"Could not focus window for export: {e}")
            except Exception as e:
                logger.debug(f"Could not find CAD window for focusing: {e}")

            # Use ESC ESC to clear any pending commands
            document.SendCommand("\x1b\x1b")
            time.sleep(0.1)

            # Disable file dialog
            document.SendCommand("FILEDIA 0\n")
            time.sleep(0.1)

            # Execute PNGOUT command with file path
            # Sequence: command, path, selection (all), finish selection
            safe_path = self._sanitize_command_input(filepath_cad)
            document.SendCommand(f"_PNGOUT\n{safe_path}\n_ALL\n\n")

            # Wait briefly for file to be written
            time.sleep(1.5)  # Increased wait for render

            # Re-enable file dialog
            document.SendCommand("FILEDIA 1\n")

            # Verify file was created
            if not os.path.exists(filepath):
                # Try fallback: maybe it didn't like _ALL, try just \n\n
                logger.debug(
                    "PNGOUT with _ALL failed, trying with default selection..."
                )
                safe_path_fallback = self._sanitize_command_input(filepath_cad)
                document.SendCommand(f"_PNGOUT\n{safe_path_fallback}\n\n")
                time.sleep(1.5)

            if not os.path.exists(filepath):
                raise Exception(f"Export file was not created at {filepath}")

            # Convert to base64
            with open(filepath, "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode("utf-8")

            logger.info(f"View exported to {filepath}")

            return {"path": filepath, "data": encoded_string}

        except Exception as e:
            logger.error(f"Export view failed: {e}")
            raise Exception(f"Failed to export view: {e}")

    def zoom_extents(self) -> bool:
        """Zoom the active viewport to fit all drawing entities via COM.

        Returns:
            True if successful, False otherwise.
        """
        try:
            application = self._get_application("zoom_extents")
            application.ZoomExtents()
            logger.debug("Zoomed to extents")
            return True
        except Exception as e:
            logger.error(f"Failed to zoom extents: {e}")
            return False

    def refresh_view(self) -> bool:
        """Refresh the view using multiple techniques for maximum compatibility.

        Uses a combination of techniques in fallback order:
        1. Application.Refresh() (COM API - no undo/redo impact)
        2. SendCommand with REDRAW (most reliable visual update)
        3. Window click simulation (forces UI update)

        Note: REDRAW command is not wrapped in UNDO to avoid complicating
        the undo/redo stack. If refresh_view is called during user operations,
        the REDRAW will be undone by the user's undo command anyway.

        Returns:
            True if refresh was attempted (best effort approach)
        """
        try:
            application = self._get_application("refresh_view")
            document = self._get_document("refresh_view")

            # Technique 1: COM API Refresh (doesn't affect undo/redo)
            try:
                application.Refresh()
                logger.debug("Refresh: COM Refresh executed")
            except Exception as e:
                logger.debug(f"COM Refresh failed: {e}")

            # Technique 2: Send REDRAW command (most reliable visual update)
            try:
                document.SendCommand("_redraw\n")
                logger.debug("Refresh: REDRAW command sent")
            except Exception as e:
                logger.debug(f"REDRAW command failed: {e}")

            # Technique 3: Simulate click on CAD window (forces UI update)
            self._simulate_autocad_click()

            return True
        except Exception as e:
            logger.debug(f"refresh_view error: {e}")
            return False

    def undo(self, count: int = 1) -> bool:
        """Undo last action(s).

        Args:
            count: Number of operations to undo (default: 1)

        Returns:
            True if successful, False otherwise
        """
        try:
            self._validate_connection()
            if count < 1:
                logger.warning(f"Invalid undo count: {count}. Must be >= 1")
                return False

            app = self._get_application("undo")
            app.ActiveDocument.SendCommand(f"_undo {count}\n")
            logger.info(f"Undo executed ({count} operation(s))")
            return True
        except Exception as e:
            logger.error(f"Failed to undo: {e}")
            return False

    def redo(self, count: int = 1) -> bool:
        """Redo last undone action(s).

        Args:
            count: Number of operations to redo (default: 1)

        Returns:
            True if successful, False otherwise
        """
        try:
            self._validate_connection()
            if count < 1:
                logger.warning(f"Invalid redo count: {count}. Must be >= 1")
                return False

            app = self._get_application("redo")
            app.ActiveDocument.SendCommand(f"_redo {count}\n")
            logger.info(f"Redo executed ({count} operation(s))")
            return True
        except Exception as e:
            logger.error(f"Failed to redo: {e}")
            return False
