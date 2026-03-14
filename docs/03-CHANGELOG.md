# Changelog

All notable changes to multiCAD-mcp will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2026-03-14

### Security (CRITICAL)

- **Path Traversal Prevention**: Added `_validate_export_path()` to prevent directory traversal attacks in file export operations.
- **Command Injection Mitigation**: Added `_sanitize_command_input()` to sanitize CAD command inputs, preventing malicious command injection.
- **Thread-Safe Singletons**: Implemented double-checked locking pattern in `AdapterRegistry` and `ConfigManager` for thread-safe operation.
- **COM Initialization Safety**: Improved error handling in `connection_mixin.py` for COM initialization across threads.

### Added

- **Block attribute management**: `get_attrs` and `set_attrs` actions in `manage_blocks` — read and write attribute tag values on block references.
- **Modern packaging**: `pyproject.toml` with full project metadata, dev/docs dependency groups, `[tool.ruff]`, `[tool.mypy]`, `[tool.interrogate]`, and `[tool.pytest]` configuration.
- **MkDocs documentation site**: Material theme with auto-generated API reference via mkdocstrings.

### Changed

- **Unified tool architecture**: 55 specific CAD commands replaced by 7 unified dispatch tools using compact shorthand format (~85% token reduction).
  - `manage_session` (11 actions), `draw_entities` (10 types), `manage_blocks` (6 actions), `manage_layers` (9 actions), `manage_files` (5 actions), `manage_entities` (10 actions), `export_data` (4 combinations).
- **Auto-named exports**: Excel export defaults to `[drawing_name]_data.xlsx` instead of `drawing_data.xlsx`.
- **Excel improvements**: autofilter enabled on all sheets (Entities, Layers, Blocks); `limit=0` ensures full export.
- **Dashboard refactor**: removed background refresher thread; export and refresh run directly on MCP thread; centralized configuration in `config.json` (port 6666).
- **Test suite**: expanded from 62 to 171 tests.

### Performance

- **O(n*m) → O(1) Optimization**: Optimized entity lookup in `set_entities_color_bylayer()` using `HandleToObject()` API.
  - Replaced inefficient nested loop iteration with direct handle-to-object lookups.
  - Expected 60%+ improvement on drawings with 10,000+ entities.

### Bug Fixes

- **Missing Return Statement**: Fixed `_paste()` function missing return value in `entities.py`.
- **Hardcoded Version**: Updated `web/api.py` to import version from `__version__.py` instead of hardcoding.
- **JSON Error Handling**: Added JSON error handling in `_set_color_bylayer()` with proper error messages.
- **Coordinate Validation**: Improved coordinate parsing with better error messages in paste operations.

---

## [0.1.3] - 2026-02-12

### Changed - Mixin Architecture Refactor

Major refactoring of the adapter layer for better maintainability.

#### Architecture

- **Mixin-based adapter**: `autocad_adapter.py` reduced from 3,198 to 99 lines.
- **11 specialized mixins**: Each mixin handles a specific responsibility (Utility, Connection, Drawing, Layer, File, View, Selection, Entity, Manipulation, Block, Export).
- **AdapterRegistry**: Encapsulated global state in singleton class.
- **Removed NLP**: Natural language processor removed (use direct tool calls).

#### Bug Fixes

- Fixed `@staticmethod` error in `validate_lineweight`.

#### Improvements

- **Refactored `DrawingMixin`**: Reduces boilerplate code in drawing methods using `_finalize_entity` helper.
- **Simplified `draw_mleader`**: Extracted complex fallback logic to improve readability.
- **Documentation**: Simplified `README.md` and updated documentation structure.

---

## [0.1.2] - 2025-12-09

### Added

- **Block creation**: `create_block` tool (from handles or selection).
- Core methods: `create_block_from_entities()`, `create_block_from_selection()`.
- 7 new tests (42 total).

### Changed

- Direct instantiation: `AutoCADAdapter(cad_type)` replaces factory.
- Context managers: `com_session()`, `SelectionSetManager`.
- Performance: `PickfirstSelectionSet` for fast entity access.

---

## [0.1.1] - 2025-11-22

### Added - Batch Operations

**13 batch operation tools** (legacy tools replaced by current unified architecture in 0.2.0):
- Drawing: `draw_lines`, `draw_circles`, `draw_arcs`, `draw_rectangles`, `draw_polylines`, `draw_texts`, `add_dimensions`.
- Layers: `rename_layers`, `delete_layers`, `turn_layers_on`, `turn_layers_off`.
- Entities: `change_entities_colors`, `change_entities_layers`.

---

## [0.1.0] - 2025-11-12

### Initial Release

- **Multi-CAD support**: AutoCAD, ZWCAD, GstarCAD, BricsCAD.
- **FastMCP 2.0** server with MCP tools.
- **Universal adapter** via COM API.
- **Excel export** with locale support.
- **Type safety**: 100% type hints.
- **Testing**: Comprehensive test suite.
