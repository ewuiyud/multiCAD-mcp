"""
Entity mixin for AutoCAD adapter.

Handles entity property operations.
"""

import logging
from typing import Dict, Any, TYPE_CHECKING

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
