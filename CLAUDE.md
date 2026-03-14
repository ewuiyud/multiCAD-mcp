# CLAUDE.md

MCP server for controlling CAD apps (AutoCAD, ZWCAD, GstarCAD, BricsCAD) via Windows COM/pywin32.
**Stack**: Python 3.10+ | FastMCP 2.0 | uv | Windows-only

---

## Commands

```powershell
uv run python src/server.py                                    # Run server
uv run pytest tests/ -v                                        # Tests (181 unit, manual excluded)
uv run ruff check src/ && uv run ruff format src/              # Lint + format
uv run mypy src/                                               # Type check (config: mypy.ini)
npx -y @modelcontextprotocol/inspector uv run python src/server.py  # MCP Inspector
```

---

## Architecture

Three layers: `server.py` → `mcp_tools/tools/` (7 unified tools) → `adapters/` (mixin-based COM adapter).

**7 tools / 55 commands** — each tool dispatches multiple actions via shorthand (`action|param|param`):
`manage_session` (11) | `draw_entities` (10) | `manage_layers` (9) | `manage_blocks` (6) | `manage_files` (5) | `manage_entities` (10) | `export_data` (4)

**Adapter composition** — `AutoCADAdapter` inherits 11 mixins + `CADInterface` ABC. All 4 CAD types share one adapter (identical COM API). Adding a new CAD type = add entry to `src/config.json` only.

**`_archive/`** (root) — deprecated backup code, excluded from all analysis, tests, and type checks.

---

## Critical Rules

### Imports — absolute only
```python
from core import CADInterface          # ✓
from ..core import CADInterface        # ✗ — breaks server.py sys.path setup
```

### Coordinates
- Input: 2D `(x,y)` or 3D `(x,y,z)` — always normalize via `normalize_coordinate()` → internal 3D
- To COM: always via `_to_variant_array()` — never pass tuples directly to COM calls

### Excel export
```python
cell.value = round(length, 3)   # Float, NOT string — regional decimal separator
cell.number_format = '0.000'
```

### COLOR_MAP (`src/mcp_tools/constants.py`)
`{"red": 1, "blue": 5, "white": 7, ...}` — **do not modify** (AutoCAD Color Index, fixed standard)

### Exception constructors
`CADConnectionError(cad_type: str, reason: str)` — **two required args**, not one string.

### Batch drawing — skip intermediate refreshes
```python
adapter.draw_line(..., _skip_refresh=True)  # In loops
adapter.refresh_view()                       # Once after loop
```

### Dashboard
Port configured in `src/config.json` → `dashboard.port` (default: **6666**).

---

## Adding a New Operation (recipe)

1. Abstract method → `src/core/cad_interface.py`
2. Implementation → relevant mixin in `src/adapters/mixins/`
3. Register in matching tool → `src/mcp_tools/tools/` (dispatch table + shorthand parser if new action)
4. Test → `tests/unit/`

---

## Workflow

### Before committing
```powershell
uv run pytest tests/ -v     # must pass
uv run ruff format src/     # format
uv run mypy src/            # must be clean
```

### Commit policy
⚠️ **Always ask the user before committing.**
Format: `git commit -m "feat(scope): description"`

### Version
Single source of truth: `src/__version__.py` — edit only this file.
