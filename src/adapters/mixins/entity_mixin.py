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
        test (ray-casting) against every TEXT/MTEXT, LINE, and POLYLINE entity in
        ModelSpace (regardless of layer) and returns results grouped by region.

        Line membership is decided by centroid: a segment belongs to a region when its
        midpoint (for LINE) or vertex centroid (for POLYLINE) lies inside the boundary.

        Also attempts automatic table parsing from the grid lines and texts.

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
                  lines (list) – LINE / POLYLINE segments whose centroid is inside;
                      polylines with >2 vertices include a "vertices" field
                  table_parse (dict) – auto-parsed table structure:
                      success (bool)
                      rows (list[list[str]]) – 2-D cell array, top-to-bottom
                      n_rows / n_cols (int)
                      unmatched_texts (list[str]) – texts outside the detected grid
        """
        import math
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

        def _cluster(values: List[float], tol: float) -> List[float]:
            """Group close float values and return cluster centres (sorted ascending)."""
            if not values:
                return []
            sorted_v = sorted(values)
            clusters: List[List[float]] = [[sorted_v[0]]]
            for v in sorted_v[1:]:
                if v - clusters[-1][-1] <= tol:
                    clusters[-1].append(v)
                else:
                    clusters.append([v])
            return [sum(c) / len(c) for c in clusters]

        def _try_parse_table(
            texts: List[Dict[str, Any]], lines: List[Dict[str, Any]]
        ) -> Dict[str, Any]:
            """Attempt to derive a 2-D table from coordinate data.

            Decomposes every line/polyline into individual segments, classifies them
            as horizontal or vertical within a 10-degree tolerance, clusters the
            resulting row/column boundaries, and maps texts to cells.
            """
            if not texts:
                return {"success": False, "reason": "no texts in region"}

            # --- collect individual H/V segments ----------------------------
            h_segs: List[tuple] = []  # (y, x_min, x_max)
            v_segs: List[tuple] = []  # (x, y_min, y_max)

            def _classify_seg(p1: List[float], p2: List[float]) -> None:
                dx = p2[0] - p1[0]
                dy = p2[1] - p1[1]
                length = math.hypot(dx, dy)
                if length < 1e-6:
                    return
                angle = abs(math.degrees(math.atan2(abs(dy), abs(dx))))
                if angle <= 10:  # horizontal (within 10°)
                    y = (p1[1] + p2[1]) / 2
                    h_segs.append((y, min(p1[0], p2[0]), max(p1[0], p2[0])))
                elif angle >= 80:  # vertical (within 10°)
                    x = (p1[0] + p2[0]) / 2
                    v_segs.append((x, min(p1[1], p2[1]), max(p1[1], p2[1])))

            for line in lines:
                verts = line.get("vertices")
                if verts and len(verts) >= 2:
                    for i in range(len(verts) - 1):
                        _classify_seg(verts[i], verts[i + 1])
                else:
                    _classify_seg(line["start_point"], line["end_point"])

            if not h_segs or not v_segs:
                # ── fallback: infer grid from text-position gaps ───────────
                # When grid lines are absent (borders on a hidden/different layer),
                # use natural coordinate gaps between text entities to infer
                # row/column boundaries.  We look for unusually large jumps
                # between consecutive sorted X or Y values: threshold = Q1 * 5.
                def _gap_boundaries(sorted_vals: List[float],
                                    abs_min: float = 30.0) -> List[float]:
                    """Find column/row split points using bimodal gap detection.

                    CAD table text positions have two gap populations:
                    - Small gaps (within one cell, split text entities)
                    - Large gaps (between distinct columns/rows)

                    We find the largest "jump" between consecutive sorted gap
                    sizes; that inflection point separates the two populations.
                    When gaps are roughly uniform (ratio < 4), every gap is a
                    boundary (no split-text cells present).
                    """
                    if len(sorted_vals) < 2:
                        return []
                    diffs = [sorted_vals[i + 1] - sorted_vals[i]
                             for i in range(len(sorted_vals) - 1)]
                    min_d = min(diffs)
                    max_d = max(diffs)

                    # Uniform spacing → every gap is a column/row boundary
                    if max_d / (min_d + 1e-6) < 4.0:
                        return [
                            (sorted_vals[i] + sorted_vals[i + 1]) / 2
                            for i in range(len(sorted_vals) - 1)
                        ]

                    # Bimodal: find the largest jump between consecutive
                    # sorted diff values — that separates within-cell gaps
                    # from between-column gaps.
                    sd = sorted(diffs)
                    jump_sizes = [sd[i + 1] - sd[i] for i in range(len(sd) - 1)]
                    split_idx = max(range(len(jump_sizes)),
                                    key=lambda i: jump_sizes[i])
                    threshold = max(
                        (sd[split_idx] + sd[split_idx + 1]) / 2,
                        min_d * 2,
                        abs_min,
                    )
                    return [
                        (sorted_vals[i] + sorted_vals[i + 1]) / 2
                        for i, d in enumerate(diffs) if d >= threshold
                    ]

                uniq_y = sorted(set(round(t["insertion_point"][1], 1) for t in texts))
                uniq_x = sorted(set(round(t["insertion_point"][0], 1) for t in texts))

                y_breaks = _gap_boundaries(uniq_y)
                x_breaks = _gap_boundaries(uniq_x)

                if not y_breaks and not x_breaks and len(uniq_y) < 2:
                    return {
                        "success": False,
                        "reason": "insufficient axis-aligned grid lines "
                        f"(h={len(h_segs)}, v={len(v_segs)}) and too few "
                        "text positions for gap inference",
                    }

                def _bucket(val: float, breaks: List[float],
                            descending: bool = False) -> int:
                    idx = sum(1 for b in breaks if val > b)
                    return (len(breaks) - idx) if descending else idx

                n_rows_fb = len(y_breaks) + 1
                n_cols_fb = len(x_breaks) + 1

                grid_fb: List[List[List]] = [
                    [[] for _ in range(n_cols_fb)] for _ in range(n_rows_fb)
                ]
                for t in texts:
                    tx, ty = t["insertion_point"][0], t["insertion_point"][1]
                    r = _bucket(ty, y_breaks, descending=True)
                    c = _bucket(tx, x_breaks, descending=False)
                    r = max(0, min(r, n_rows_fb - 1))
                    c = max(0, min(c, n_cols_fb - 1))
                    # store (x, text) so we can sort by X before joining
                    grid_fb[r][c].append((tx, t["text_string"]))

                # Join texts within same cell sorted by X position.
                # No separator: CAD often splits one logical string ("FM丙0620")
                # into adjacent text entities — concatenation restores it.
                rows_fb = [
                    ["".join(s for _, s in sorted(cell)) for cell in row]
                    for row in grid_fb
                ]

                return {
                    "success": True,
                    "n_rows": n_rows_fb,
                    "n_cols": n_cols_fb,
                    "rows": rows_fb,
                    "unmatched_texts": [],
                    "note": "grid inferred from text-position gaps (no grid lines found)",
                }

            # --- auto tolerance: 1% of smaller table dimension --------------
            all_y = [y for y, _, _ in h_segs]
            all_x = [x for x, _, _ in v_segs]
            y_span = max(all_y) - min(all_y) if len(all_y) > 1 else 1.0
            x_span = max(all_x) - min(all_x) if len(all_x) > 1 else 1.0
            tol = max(2.0, min(y_span, x_span) * 0.01)

            row_ys = _cluster(all_y, tol)   # ascending
            col_xs = _cluster(all_x, tol)   # ascending

            if len(row_ys) < 2 or len(col_xs) < 2:
                return {
                    "success": False,
                    "reason": f"grid too sparse: {len(row_ys)} h-lines, "
                    f"{len(col_xs)} v-lines after clustering",
                }

            # CAD Y increases upward → top row = highest Y
            row_ys_desc = sorted(row_ys, reverse=True)
            n_rows = len(row_ys_desc) - 1
            n_cols = len(col_xs) - 1

            # --- assign texts to cells --------------------------------------
            grid: List[List[List[str]]] = [
                [[] for _ in range(n_cols)] for _ in range(n_rows)
            ]
            unmatched: List[str] = []

            for t in texts:
                ip = t["insertion_point"]
                tx, ty = ip[0], ip[1]

                r_idx = next(
                    (
                        r
                        for r in range(n_rows)
                        if row_ys_desc[r + 1] - tol <= ty <= row_ys_desc[r] + tol
                    ),
                    None,
                )
                c_idx = next(
                    (
                        c
                        for c in range(n_cols)
                        if col_xs[c] - tol <= tx <= col_xs[c + 1] + tol
                    ),
                    None,
                )

                if r_idx is not None and c_idx is not None:
                    grid[r_idx][c_idx].append(t["text_string"])
                else:
                    unmatched.append(t["text_string"])

            rows = [[" ".join(cell) for cell in row] for row in grid]

            return {
                "success": True,
                "n_rows": n_rows,
                "n_cols": n_cols,
                "rows": rows,
                "unmatched_texts": unmatched,
            }

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

                    # ── all polyline types ─────────────────────────────────────
                    if "POLY" in obj_upper:
                        dyn = win32com.client.dynamic.Dispatch(entity)
                        is_closed = bool(getattr(dyn, "Closed", False))
                        coords_raw = getattr(dyn, "Coordinates", None)
                        if coords_raw is None:
                            continue
                        flat = [float(v) for v in coords_raw]
                        stride = 3 if "3D" in obj_upper else 2
                        verts = [
                            flat[i : i + stride]
                            for i in range(0, len(flat), stride)
                        ]

                        if is_closed and entity_layer == target_layer:
                            # closed boundary on target layer → region
                            area_raw = getattr(dyn, "Area", None)
                            closed_polylines.append(
                                {
                                    "handle": handle,
                                    "object_type": obj_name,
                                    "coords": verts,
                                    "area": float(area_raw)
                                    if area_raw is not None
                                    else None,
                                }
                            )
                        elif len(verts) >= 2:
                            # non-closed polyline (any layer) OR closed on other layer
                            # → treat as grid / divider lines
                            cx = sum(v[0] for v in verts) / len(verts)
                            cy = sum(v[1] for v in verts) / len(verts)
                            length_raw = getattr(dyn, "Length", None)
                            entry: Dict[str, Any] = {
                                "handle": handle,
                                "layer": str(getattr(entity, "Layer", "")),
                                "start_point": verts[0],
                                "end_point": verts[-1],
                                "length": float(length_raw)
                                if length_raw is not None
                                else None,
                                "_mid": [cx, cy],
                            }
                            if len(verts) > 2:
                                entry["vertices"] = verts
                            line_entities.append(entry)

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

                    # ── single LINE entities (any layer) ──────────────────────
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
                        entry = {
                            "handle": le["handle"],
                            "layer": le["layer"],
                            "start_point": le["start_point"],
                            "end_point": le["end_point"],
                            "length": le["length"],
                        }
                        if "vertices" in le:
                            entry["vertices"] = le["vertices"]
                        matched_lines.append(entry)

                table_parse = _try_parse_table(matched_texts, matched_lines)

                regions.append(
                    {
                        "handle": pl["handle"],
                        "object_type": pl["object_type"],
                        "area": pl["area"],
                        "texts": matched_texts,
                        "lines": matched_lines,
                        "table_parse": table_parse,
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
