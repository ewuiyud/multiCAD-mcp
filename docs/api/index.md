# API Reference

Auto-generated from source docstrings. All classes and public methods are documented here.

## Modules

| Module | Description |
|--------|-------------|
| [AutoCAD Adapter](adapter.md) | Main composite adapter class |
| [Adapter Manager](adapter_manager.md) | Registry singleton for managing adapter lifecycle |
| **Mixins** | |
| [Connection](mixins/connection.md) | COM connection and lifecycle |
| [Drawing](mixins/drawing.md) | Geometric entity creation |
| [Export](mixins/export.md) | Excel export and entity data extraction |
| [Layer](mixins/layer.md) | Layer management |
| [Block](mixins/block.md) | Block definition and insertion |
| [Manipulation](mixins/manipulation.md) | Entity transform operations |
| [Selection](mixins/selection.md) | Entity selection |
| [File](mixins/file.md) | Drawing file operations |
| [View](mixins/view.md) | Viewport and undo/redo |
| [Entity](mixins/entity.md) | Entity property access |
| [Utility](mixins/utility.md) | Helpers and converters |
| [MCP Tools](tools.md) | Tool registration layer |
| [Core](core.md) | Interfaces, config, exceptions |

## Quick Start

```python
from adapters import AutoCADAdapter

# Connect to a running AutoCAD instance
adapter = AutoCADAdapter("autocad")
adapter.connect()

# Draw entities
handle = adapter.draw_line((0, 0), (100, 0), layer="0", color="white")

# Export to Excel
adapter.export_to_excel("output.xlsx")
```
