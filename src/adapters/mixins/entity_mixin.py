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

    def extract_texts_in_closed_polylines(
        self, layer_name: str
    ) -> Dict[str, Any]:
        """Find all closed polylines on a layer and extract texts and lines within each boundary.

        For each closed polyline found on ``layer_name``, performs a point-in-polygon
        test (ray-casting) against every TEXT/MTEXT and LINE entity in ModelSpace
        (regardless of layer) and returns results grouped by polyline region.

        Line membership is decided by midpoint: a line belongs to a region when its
        midpoint lies inside the closed boundary.

        Args:
            layer_name: Name of the layer to scan for closed polylines.

        Returns:
            Dict with keys:
              success (bool)
              layer (str)
              total_regions (int)
              regions (list) – one entry per closed polyline, each has:
                  handle (str)
                  object_type (str)
                  area (float | None)
                  texts (list) – texts whose insertion point is inside
                  lines (list) – line segments whose midpoint is inside
        """
        import win32com.client

        def _point_in_polygon(px: float, py: float, polygon: List[List[float]]) -> bool:
            """Ray-casting algorithm for 2-D point-in-polygon test."""
            n = len(polygon)
            if n < 3:
                return False
            inside = False
            x, y = px, py
            j = n - 1
            for i in range(n):
                xi, yi = polygon[i][0], polygon[i][1]
                xj, yj = polygon[j][0], polygon[j][1]
                if ((yi > y) != (yj > y)) and (
                    x < (xj - xi) * (y - yi) / (yj - yi + 1e-12) + xi
                ):
                    inside = not inside
                j = i
            return inside

        try:
            document = self._get_document("extract_texts_in_closed_polylines")
            target_layer = layer_name.strip().lower()

            closed_polylines: List[Dict[str, Any]] = []
            text_entities: List[Dict[str, Any]] = []
            line_entities: List[Dict[str, Any]] = []

            for entity in document.ModelSpace:
                try:
                    obj_name = str(getattr(entity, "ObjectName", ""))
                    entity_layer = str(getattr(entity, "Layer", "")).strip().lower()
                    handle = str(entity.Handle)
                    obj_upper = obj_name.upper()

                    # ── closed polylines on target layer ──────────────────────
                    if "POLY" in obj_upper and entity_layer == target_layer:
                        dyn = win32com.client.dynamic.Dispatch(entity)
                        if getattr(dyn, "Closed", False):
                            coords_raw = getattr(dyn, "Coordinates", None)
                            if coords_raw is not None:
                                flat = [float(v) for v in coords_raw]
                                stride = 3 if "3D" in obj_upper else 2
                                coords = [
                                    flat[i : i + stride]
                                    for i in range(0, len(flat), stride)
                                ]
                                area_raw = getattr(dyn, "Area", None)
                                closed_polylines.append(
                                    {
                                        "handle": handle,
                                        "object_type": obj_name,
                                        "coords": coords,
                                        "area": float(area_raw)
                                        if area_raw is not None
                                        else None,
                                    }
                                )

                    # ── text / mtext entities (any layer) ─────────────────────
                    elif "TEXT" in obj_upper:
                        dyn = win32com.client.dynamic.Dispatch(entity)
                        ip_raw = getattr(dyn, "InsertionPoint", None)
                        if ip_raw is not None:
                            ip = [float(v) for v in ip_raw]
                            text_str = getattr(dyn, "TextString", None) or getattr(
                                dyn, "Contents", None
                            )
                            text_entities.append(
                                {
                                    "handle": handle,
                                    "layer": str(getattr(entity, "Layer", "")),
                                    "text_string": str(text_str)
                                    if text_str is not None
                                    else "",
                                    "insertion_point": ip,
                                }
                            )

                    # ── line entities (any layer) ──────────────────────────────
                    elif obj_upper == "ACDBLINE":
                        dyn = win32com.client.dynamic.Dispatch(entity)
                        sp_raw = getattr(dyn, "StartPoint", None)
                        ep_raw = getattr(dyn, "EndPoint", None)
                        if sp_raw is not None and ep_raw is not None:
                            sp = [float(v) for v in sp_raw]
                            ep = [float(v) for v in ep_raw]
                            length_raw = getattr(dyn, "Length", None)
                            line_entities.append(
                                {
                                    "handle": handle,
                                    "layer": str(getattr(entity, "Layer", "")),
                                    "start_point": sp,
                                    "end_point": ep,
                                    "length": float(length_raw)
                                    if length_raw is not None
                                    else None,
                                    # midpoint used for containment test
                                    "_mid": [
                                        (sp[0] + ep[0]) / 2,
                                        (sp[1] + ep[1]) / 2,
                                    ],
                                }
                            )

                except Exception as e:
                    logger.debug(f"Skipped entity during scan: {e}")
                    continue

            # ── match texts and lines to each region ──────────────────────────
            regions = []
            for pl in closed_polylines:
                polygon = pl["coords"]

                matched_texts = []
                for te in text_entities:
                    ip = te["insertion_point"]
                    if _point_in_polygon(ip[0], ip[1], polygon):
                        matched_texts.append(
                            {
                                "handle": te["handle"],
                                "layer": te["layer"],
                                "text_string": te["text_string"],
                                "insertion_point": te["insertion_point"],
                            }
                        )

                matched_lines = []
                for le in line_entities:
                    mid = le["_mid"]
                    if _point_in_polygon(mid[0], mid[1], polygon):
                        matched_lines.append(
                            {
                                "handle": le["handle"],
                                "layer": le["layer"],
                                "start_point": le["start_point"],
                                "end_point": le["end_point"],
                                "length": le["length"],
                            }
                        )

                regions.append(
                    {
                        "handle": pl["handle"],
                        "object_type": pl["object_type"],
                        "area": pl["area"],
                        "texts": matched_texts,
                        "lines": matched_lines,
                    }
                )

            return {
                "success": True,
                "layer": layer_name,
                "total_regions": len(regions),
                "regions": regions,
            }

        except Exception as e:
            logger.error(f"extract_texts_in_closed_polylines failed: {e}")
            return {"success": False, "layer": layer_name, "error": str(e)}

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
