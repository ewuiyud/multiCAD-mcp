"""
Entity mixin for AutoCAD adapter.

Handles entity property operations.
"""

import logging
from typing import Dict, Any, List, Optional, TYPE_CHECKING

logger = logging.getLogger(__name__)


class EntityMixin:
    """Mixin for entity property operations."""

    if TYPE_CHECKING:

        def _get_document(self, operation: str = "operation") -> Any: ...
        def _get_color_index(self, color_name: str) -> int: ...
        def validate_lineweight(self, weight: int) -> int: ...

    def delete_entity(self, handle: str) -> bool:
        """Delete a drawing entity identified by its COM handle.

        Args:
            handle: The entity handle string as returned by AutoCAD COM.

        Returns:
            True if deleted successfully, False otherwise.
        """
        try:
            document = self._get_document("delete_entity")

            entity = document.HandleToObject(handle)
            entity.Delete()
            logger.debug(f"Deleted entity {handle}")
            return True
        except Exception as e:
            logger.error(f"Failed to delete entity: {e}")
            return False

    def get_entity_properties(self, handle: str) -> Dict[str, Any]:
        """Retrieve common properties of a drawing entity via COM.

        Args:
            handle: The entity handle string as returned by AutoCAD COM.

        Returns:
            Dictionary with keys: handle, object_name, layer, color, lineweight.
            Returns an empty dict if the entity cannot be accessed.
        """
        try:
            document = self._get_document("get_entity_properties")

            entity = document.HandleToObject(handle)
            return {
                "handle": entity.Handle,
                "object_name": entity.ObjectName,
                "layer": entity.Layer,
                "color": entity.Color,
                "lineweight": entity.LineWeight,
            }
        except Exception as e:
            logger.error(f"Failed to get entity properties: {e}")
            return {}

    def query_entity_geometry(self, handle: str) -> Dict[str, Any]:
        """Query detailed geometry of a single entity by handle.

        Returns common properties plus type-specific geometry:
        - LINE: start_point, end_point, length
        - CIRCLE: center, radius
        - ARC: center, radius, start_angle, end_angle, length
        - LWPOLYLINE/POLYLINE: coordinates, length, area, is_closed
        - SPLINE: coordinates, length
        - TEXT/MTEXT: insertion_point, text_string
        - INSERT (block ref): insertion_point, name

        Args:
            handle: The entity handle string as returned by AutoCAD COM.

        Returns:
            Dict with success, handle, object_type, layer, and geometry fields.
            Returns success=False with error on failure.
        """
        import win32com.client

        def _to_list(val) -> Optional[List[float]]:
            """Convert COM tuple/array to plain Python list, or None."""
            if val is None:
                return None
            try:
                return [float(v) for v in val]
            except (TypeError, ValueError):
                return None

        try:
            document = self._get_document("query_entity_geometry")
            entity = document.HandleToObject(handle)
            dyn = win32com.client.dynamic.Dispatch(entity)

            object_type = str(getattr(entity, "ObjectName", "Unknown"))
            layer = str(getattr(entity, "Layer", ""))
            color_raw = getattr(entity, "Color", 256)
            color = int(color_raw.ColorIndex) if hasattr(color_raw, "ColorIndex") else int(color_raw)
            type_upper = object_type.upper()

            result: Dict[str, Any] = {
                "success": True,
                "handle": handle,
                "object_type": object_type,
                "layer": layer,
                "color": color,
            }

            if "LINE" in type_upper and "POLY" not in type_upper and "SPLINE" not in type_upper:
                sp = _to_list(getattr(dyn, "StartPoint", None))
                ep = _to_list(getattr(dyn, "EndPoint", None))
                length = getattr(dyn, "Length", None)
                result["start_point"] = sp
                result["end_point"] = ep
                result["length"] = float(length) if length is not None else None

            elif "CIRCLE" in type_upper:
                center = _to_list(getattr(dyn, "Center", None))
                radius = getattr(dyn, "Radius", None)
                result["center"] = center
                result["radius"] = float(radius) if radius is not None else None

            elif "ARC" in type_upper:
                center = _to_list(getattr(dyn, "Center", None))
                radius = getattr(dyn, "Radius", None)
                start_angle = getattr(dyn, "StartAngle", None)
                end_angle = getattr(dyn, "EndAngle", None)
                length = getattr(dyn, "ArcLength", None) or getattr(dyn, "Length", None)
                result["center"] = center
                result["radius"] = float(radius) if radius is not None else None
                result["start_angle"] = float(start_angle) if start_angle is not None else None
                result["end_angle"] = float(end_angle) if end_angle is not None else None
                result["length"] = float(length) if length is not None else None

            elif "POLY" in type_upper or "SPLINE" in type_upper:
                coords_raw = getattr(dyn, "Coordinates", None)
                coords: Optional[List] = None
                if coords_raw is not None:
                    try:
                        flat = [float(v) for v in coords_raw]
                        # LWPOLYLINE returns flat [x0,y0,x1,y1,...]; POLYLINE may return [x,y,z,...]
                        stride = 3 if "SPLINE" in type_upper or "3D" in type_upper else 2
                        coords = [flat[i:i + stride] for i in range(0, len(flat), stride)]
                    except (TypeError, ValueError):
                        coords = None
                length = getattr(dyn, "Length", None) or getattr(dyn, "TotalLength", None)
                area = getattr(dyn, "Area", None)
                is_closed = getattr(dyn, "Closed", None)
                result["coordinates"] = coords
                result["length"] = float(length) if length is not None else None
                result["area"] = float(area) if area is not None else None
                result["is_closed"] = bool(is_closed) if is_closed is not None else None

            elif "TEXT" in type_upper:
                ip = _to_list(getattr(dyn, "InsertionPoint", None))
                text = getattr(dyn, "TextString", None) or getattr(dyn, "Contents", None)
                result["insertion_point"] = ip
                result["text_string"] = str(text) if text is not None else None

            elif "INSERT" in type_upper:
                ip = _to_list(getattr(dyn, "InsertionPoint", None))
                name = getattr(dyn, "Name", None)
                result["insertion_point"] = ip
                result["name"] = str(name) if name is not None else None

            return result

        except Exception as e:
            logger.error(f"Failed to query entity geometry for handle '{handle}': {e}")
            return {"success": False, "handle": handle, "error": str(e)}

    def set_entity_properties(self, handle: str, properties: Dict[str, Any]) -> bool:
        """Modify one or more properties of a drawing entity via COM.

        Supported keys in ``properties``: ``layer`` (str), ``color`` (str or int ACI
        index), ``lineweight`` (int).

        Args:
            handle: The entity handle string as returned by AutoCAD COM.
            properties: Dictionary of property names to new values.

        Returns:
            True if the update succeeded, False otherwise.
        """
        try:
            document = self._get_document("set_entity_properties")

            entity = document.HandleToObject(handle)

            if "layer" in properties:
                entity.Layer = properties["layer"]
            if "color" in properties:
                color = properties["color"]
                if isinstance(color, str):
                    color = self._get_color_index(color)
                entity.Color = color
            if "lineweight" in properties:
                entity.LineWeight = self.validate_lineweight(properties["lineweight"])

            logger.debug(f"Updated properties for entity {handle}")
            return True
        except Exception as e:
            logger.error(f"Failed to set entity properties: {e}")
            return False
