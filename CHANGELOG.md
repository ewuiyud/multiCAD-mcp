# Changelog

All notable changes to multiCAD-mcp will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2025-02-16

### Security (CRITICAL)

- **Path Traversal Prevention**: Added `_validate_export_path()` to prevent directory traversal attacks in file export operations
- **Command Injection Mitigation**: Added `_sanitize_command_input()` to sanitize CAD command inputs, preventing malicious command injection
- **Thread-Safe Singletons**: Implemented double-checked locking pattern in `AdapterRegistry` and `ConfigManager` for thread-safe operation
- **COM Initialization Safety**: Improved error handling in `connection_mixin.py` for COM initialization across threads

### Features

- **Input Validation**: Added `_validate_drawing_params()` for sanity checks on drawing operations (radius, angle, points validation)
- **Robust Window Detection**: Added `_find_cad_window()` with strict window matching using both title and class name, preventing VBA editor mismatches
- **Web Dashboard**: Integrated web dashboard at `http://localhost:8080` with real-time CAD status, layer/block browser, and entity viewer
- **Configuration Constants**: Centralized magic numbers in `constants.py` for improved maintainability

### Performance

- **O(n*m) → O(1) Optimization**: Optimized entity lookup in `set_entities_color_bylayer()` using `HandleToObject()` API
  - Replaced inefficient nested loop iteration with direct handle-to-object lookups
  - Expected 60%+ improvement on drawings with 10,000+ entities

### Bug Fixes

- **Missing Return Statement**: Fixed `_paste()` function missing return value in entities.py
- **Hardcoded Version**: Updated web/api.py to import version from `__version__.py` instead of hardcoding
- **JSON Error Handling**: Added JSON error handling in `_set_color_bylayer()` with proper error messages
- **Coordinate Validation**: Improved coordinate parsing with better error messages in paste operations

### Code Quality

- **Standardized Error Handling**: Extracted `_handle_operation_error()` helper for consistent error classification across mixins
- **Eliminated Duplication**: Extracted `_iterate_entities_safe()` helper to remove duplicated entity iteration patterns
- **Code Formatting**: Applied black formatter to entire codebase for consistency
- **Documentation**: Added comprehensive web dashboard documentation and updated version references

### Testing

- **Security Test Suite**: Added 14 comprehensive security tests covering:
  - Path traversal prevention (2 tests)
  - Command injection prevention (4 tests)
  - Thread safety verification (3 tests)
  - Input validation robustness (5 tests)
- **Total Test Coverage**: 171 unit tests (14 new + 157 existing)

### Dependencies

- No new dependencies added
- All changes backward compatible with existing API

### Migration Notes

- Version bumped from 0.1.3 to 0.2.0 (MINOR version bump)
- No breaking changes - all existing MCP tools maintain compatibility
- Web dashboard is optional - server works without it

### Contributors

- Claude Haiku 4.5

---

## [0.1.3] - Previous Release

See git history for previous changes.
