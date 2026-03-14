"""
Layer mixin for AutoCAD adapter.

Handles all layer management operations.
"""

import logging
from typing import List, Dict, Any, Optional, TYPE_CHECKING

from core import LayerError, ColorError
from mcp_tools.constants import COLOR_MAP

logger = logging.getLogger(__name__)


class LayerMixin:
    """Mixin for layer management operations."""

    if TYPE_CHECKING:
        _drawing_state: dict

        def _validate_connection(self) -> None: ...

        def _get_document(self, operation: str = "operation") -> Any: ...

        def _get_application(self, operation: str = "operation") -> Any: ...

        def _get_color_index(self, color_name: str) -> int: ...

        def validate_lineweight(self, weight: int) -> int: ...

        def _fast_get_property(
            self, obj: Any, property_name: str, default: Any = None
        ) -> Any: ...

        def _safe_get_property(
            self, obj: Any, property_name: str, default: Any = None
        ) -> Any: ...

    def create_layer(
        self,
        name: str,
        color: str | int = "white",
        lineweight: int = 0,
    ) -> bool:
        """Create a new layer in the active drawing via COM.

        Args:
            name: Name for the new layer.
            color: Layer color as a name (e.g. ``"red"``) or ACI index. Default: ``"white"``.
            lineweight: Line weight in hundredths of a millimetre (e.g. 25 = 0.25 mm).
                Use 0 for the default line weight.

        Returns:
            True if the layer was created successfully, False otherwise.
        """
        try:
            document = self._get_document("create_layer")

            layer_obj = document.Layers.Add(name)

            if isinstance(color, str):
                color = self._get_color_index(color)
            layer_obj.Color = color

            if self.validate_lineweight(lineweight) == lineweight:
                layer_obj.LineWeight = lineweight

            logger.info(f"Created layer '{name}'")
            return True

        except Exception as e:
            logger.error(f"Failed to create layer '{name}': {e}")
            return False

    def set_current_layer(self, name: str) -> bool:
        """Set the active (current) drawing layer via COM.

        New entities will be created on this layer by default.

        Args:
            name: Name of an existing layer to make current.

        Returns:
            True if successful, False otherwise.
        """
        try:
            document = self._get_document("set_current_layer")

            document.ActiveLayer = document.Layers.Item(name)
            self._drawing_state["current_layer"] = name
            logger.debug(f"Set current layer to '{name}'")
            return True
        except Exception as e:
            logger.error(f"Failed to set current layer: {e}")
            return False

    def get_current_layer(self) -> str:
        """Return the name of the currently active layer.

        Falls back to the cached drawing state if the COM call fails.

        Returns:
            Active layer name string, or ``"0"`` if undetermined.
        """
        try:
            document = self._get_document("get_current_layer")
            return str(document.ActiveLayer.Name)
        except Exception:
            current_layer = self._drawing_state["current_layer"]
            return str(current_layer) if current_layer else "0"

    def list_layers(self) -> List[str]:
        """Return the names of all layers in the active drawing via COM.

        Returns:
            List of layer name strings. Empty list on error.
        """
        try:
            document = self._get_document("list_layers")
            layers: List[str] = []
            for layer in document.Layers:
                layers.append(layer.Name)
            return layers
        except Exception as e:
            logger.error(f"Failed to list layers: {e}")
            return []

    def get_layers_info(
        self, entity_data: Optional[List[Dict[str, Any]]] = None
    ) -> List[Dict[str, Any]]:
        """Get detailed information about all layers.

        Optimized to count entities per layer in a single pass using direct iteration,
        or from pre-extracted entity data to avoid re-iterating ModelSpace.

        Args:
            entity_data: Optional pre-extracted entity data. If provided, layer counts
                        will be computed from this data instead of iterating ModelSpace.

        Returns:
            List of dictionaries with layer information:
            - Name: Layer name
            - ObjectCount: Number of objects on the layer
            - Color: Layer color
            - Linetype: Layer linetype
            - Lineweight: Layer lineweight
            - IsLocked: Whether layer is locked
            - IsVisible: Whether layer is visible
        """
        try:
            document = self._get_document("get_layers_info")
            layers_info = []

            # OPTIMIZATION: Use pre-extracted data if available to avoid re-iteration
            layer_counts: Dict[str, int] = {}

            if entity_data is not None:
                # Count from pre-extracted data (MUCH faster - no COM calls)
                logger.debug(
                    f"Computing layer counts from {len(entity_data)} pre-extracted entities"
                )
                for entity in entity_data:
                    layer_name = entity.get("Layer", "0")
                    layer_counts[layer_name] = layer_counts.get(layer_name, 0) + 1
            else:
                # Fallback: Use SelectionSets to quickly count entities per layer (O(K))
                logger.debug("Using SelectionSets to count entities by layer")
                import pythoncom
                import win32com.client
                import time

                perf_start = time.perf_counter()

                # Setup selection set manager helper
                from contextlib import contextmanager

                @contextmanager
                def _temp_ss(doc, name):
                    try:
                        doc.SelectionSets.Item(name).Delete()
                    except Exception:
                        pass
                    ss = doc.SelectionSets.Add(name)
                    try:
                        yield ss
                    finally:
                        try:
                            ss.Delete()
                        except Exception:
                            pass

                def to_variant_array(types, values):
                    return win32com.client.VARIANT(types, values)

                ft = to_variant_array(
                    pythoncom.VT_ARRAY | pythoncom.VT_I2, [8]
                )  # DXF Code 8: Layer Name

                with _temp_ss(document, "MCP_LAYER_COUNTS") as ss:
                    for layer in document.Layers:
                        lname = layer.Name
                        fd = to_variant_array(
                            pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [lname]
                        )
                        try:
                            ss.Clear()
                            ss.Select(5, None, None, ft, fd)  # 5 = acSelectionSetAll
                            count = ss.Count
                            if count > 0:
                                layer_counts[lname] = count
                        except Exception as e:
                            logger.debug(
                                f"Failed to count entities on layer {lname}: {e}"
                            )

                elapsed = time.perf_counter() - perf_start
                logger.info(
                    f"[PERF] Layer counting via SelectionSets took {elapsed:.3f}s"
                )

            # Build layer information
            for layer in document.Layers:
                try:
                    # Get layer properties using dynamic dispatch for robustness
                    import win32com.client

                    dyn_layer = win32com.client.dynamic.Dispatch(layer)

                    layer_color_val = 7  # Default white
                    try:
                        # Try TrueColor first for modern CAD compatibility
                        if hasattr(dyn_layer, "TrueColor"):
                            tc = dyn_layer.TrueColor
                            layer_color_val = int(getattr(tc, "ColorIndex", 7))
                        else:
                            layer_color_val = int(getattr(dyn_layer, "Color", 7))
                    except (TypeError, ValueError, AttributeError):
                        layer_color_val = 7

                    # Convert to name if possible, or keep as string numeric
                    color_map_reverse = {v: k for k, v in COLOR_MAP.items()}
                    color_name = color_map_reverse.get(
                        layer_color_val, str(layer_color_val)
                    )

                    logger.info(
                        f"Layer '{layer.Name}' extracted color: {layer_color_val} ({color_name})"
                    )

                    layer_info = {
                        "Name": str(layer.Name).strip(),
                        "ObjectCount": layer_counts.get(str(layer.Name).strip(), 0),
                        "Color": color_name,
                        "Linetype": str(
                            self._safe_get_property(layer, "Linetype", "Continuous")
                        ),
                        "Lineweight": str(
                            self._safe_get_property(layer, "Lineweight", "Default")
                        ),
                        "IsLocked": bool(self._safe_get_property(layer, "Lock", False)),
                        "IsVisible": bool(
                            self._safe_get_property(layer, "LayerOn", True)
                        )
                        and not bool(self._safe_get_property(layer, "Frozen", False)),
                    }
                    layers_info.append(layer_info)
                except Exception as e:
                    logger.debug(f"Failed to get info for layer {layer.Name}: {e}")
                    continue

            return layers_info
        except Exception as e:
            logger.error(f"Failed to get layers info: {e}")
            return []

    def rename_layer(self, old_name: str, new_name: str) -> bool:
        """Rename an existing layer via COM.

        Layer ``"0"`` cannot be renamed.

        Args:
            old_name: Current name of the layer to rename.
            new_name: New name for the layer.

        Returns:
            True if renamed successfully, False otherwise.
        """
        try:
            self._validate_connection()
            document = self._get_document("rename_layer")

            if old_name == "0":
                logger.error("Cannot rename layer '0' (standard layer)")
                return False

            layer = document.Layers.Item(old_name)
            layer.Name = new_name
            logger.info(f"Renamed layer '{old_name}' to '{new_name}'")
            return True
        except Exception as e:
            logger.error(f"Failed to rename layer '{old_name}' to '{new_name}': {e}")
            return False

    def delete_layer(self, name: str) -> bool:
        """Delete a layer from the active drawing via COM.

        Layer ``"0"`` cannot be deleted.

        Args:
            name: Name of the layer to delete.

        Returns:
            True if deleted successfully, False otherwise.
        """
        try:
            self._validate_connection()
            document = self._get_document("delete_layer")

            if name == "0":
                logger.error("Cannot delete layer '0' (standard layer)")
                return False

            layer = document.Layers.Item(name)
            layer.Delete()
            logger.info(f"Deleted layer '{name}'")
            return True
        except Exception as e:
            logger.error(f"Failed to delete layer '{name}': {e}")
            return False

    def turn_layer_on(self, name: str) -> bool:
        """Make a frozen layer visible by unfreezing it via COM.

        Args:
            name: Name of the layer to turn on.

        Returns:
            True if successful, False otherwise.
        """
        try:
            self._validate_connection()
            document = self._get_document("turn_layer_on")

            layer = document.Layers.Item(name)
            layer.Freeze = False
            logger.info(f"Turned on layer '{name}'")
            return True
        except Exception as e:
            logger.error(f"Failed to turn on layer '{name}': {e}")
            return False

    def turn_layer_off(self, name: str) -> bool:
        """Hide a layer by freezing it via COM.

        Args:
            name: Name of the layer to turn off.

        Returns:
            True if successful, False otherwise.
        """
        try:
            self._validate_connection()
            document = self._get_document("turn_layer_off")

            layer = document.Layers.Item(name)
            layer.Freeze = True
            logger.info(f"Turned off layer '{name}'")
            return True
        except Exception as e:
            logger.error(f"Failed to turn off layer '{name}': {e}")
            return False

    def is_layer_on(self, name: str) -> bool:
        """Check whether a layer is currently visible (not frozen) via COM.

        Args:
            name: Name of the layer to check.

        Returns:
            True if the layer is visible, False if frozen or on error.
        """
        try:
            self._validate_connection()
            document = self._get_document("is_layer_on")

            layer = document.Layers.Item(name)
            return not layer.Freeze
        except Exception as e:
            logger.error(f"Failed to check layer '{name}' visibility: {e}")
            return False

    def set_layer_color(self, layer_name: str, color: str | int) -> bool:
        """Set the color of a layer.

        Args:
            layer_name: Name of the layer to modify
            color: Color name (from COLOR_MAP) or ACI index (1-255)

        Returns:
            bool: True if successful, False otherwise

        Note:
            - Uses the modern TrueColor property (recommended by Autodesk)
            - Accepts color names: "red", "blue", "green", etc.
            - Accepts ACI index: 1-255
            - Color value 256 (bylayer) is not valid for layers
        """
        try:
            self._validate_connection()
            document = self._get_document("set_layer_color")
            app = self._get_application("set_layer_color")

            # Get the layer
            try:
                layer = document.Layers.Item(layer_name)
            except Exception:
                raise LayerError(layer_name, "Layer does not exist")

            # Convert color name to ACI index
            if isinstance(color, str):
                color_index = self._get_color_index(color)
            else:
                color_index = color

            # Validate color index (1-255 for layers, 256 is not valid)
            if color_index == 256:
                raise ColorError(
                    str(color_index),
                    "Color 'bylayer' (256) is not valid for layers. Use a specific ACI color (1-255).",
                )
            if not (0 <= color_index <= 255):
                raise ColorError(
                    str(color_index), "Invalid color index. Must be 0-255."
                )

            # Create AcCmColor object (modern method)
            color_obj = app.GetInterfaceObject("AutoCAD.AcCmColor.20")
            color_obj.ColorIndex = color_index

            # Apply to layer using TrueColor property
            layer.TrueColor = color_obj

            logger.info(f"Set layer '{layer_name}' color to ACI {color_index}")
            return True

        except (LayerError, ColorError):
            raise
        except Exception as e:
            logger.error(f"Failed to set layer '{layer_name}' color: {e}")
            return False

    def set_entities_color_bylayer(self, handles: List[str]) -> Dict[str, Any]:
        """Set entities to use their layer's color (ByLayer).

        Args:
            handles: List of entity handles to modify

        Returns:
            dict: Result summary with counts and details:
                - total: Total entities processed
                - changed: Number successfully changed to ByLayer
                - failed: Number that failed
                - results: List of per-entity results

        Note:
            - Assigns color value 256 (acByLayer) to entities
            - Entities will inherit their layer's color
            - Uses TrueColor property (modern method)
        """
        try:
            self._validate_connection()
            document = self._get_document("set_entities_color_bylayer")
            app = self._get_application("set_entities_color_bylayer")
            results = []
            changed_count = 0
            failed_count = 0

            # Create AcCmColor object for ByLayer (256)
            color_obj = app.GetInterfaceObject("AutoCAD.AcCmColor.20")
            color_obj.ColorIndex = 256  # acByLayer

            for handle in handles:
                try:
                    # Get entity by handle using HandleToObject (O(1) vs O(n) iteration)
                    try:
                        entity = document.HandleToObject(handle)
                    except Exception as e:
                        results.append(
                            {
                                "handle": handle,
                                "success": False,
                                "error": f"Entity not found: {e}",
                            }
                        )
                        failed_count += 1
                        continue

                    # Set color to ByLayer using TrueColor
                    entity.TrueColor = color_obj

                    results.append({"handle": handle, "success": True})
                    changed_count += 1

                except Exception as e:
                    results.append(
                        {"handle": handle, "success": False, "error": str(e)}
                    )
                    failed_count += 1

            logger.info(f"Set {changed_count}/{len(handles)} entities to ByLayer color")

            return {
                "total": len(handles),
                "changed": changed_count,
                "failed": failed_count,
                "results": results,
            }

        except Exception as e:
            logger.error(f"Failed to set entities to ByLayer: {e}")
            return {
                "total": len(handles),
                "changed": 0,
                "failed": len(handles),
                "error": str(e),
                "results": [],
            }
