# 04 - Troubleshooting

## Quick Reference

| Error | Cause | Solution |
|-------|-------|----------|
| "Not connected" | CAD not running | Start CAD application |
| "Connection failed" | COM issue | Reinstall pywin32 |
| "Invalid coordinate" | Bad format | Use "x,y" or "x,y,z" |
| "Document not available" | CAD closed | Restart CAD |
| "Permission denied" | File locked | Check permissions |

## Connection Issues

### "Connection failed: AutoCAD.Application"

```powershell
# Reinstall pywin32
uv run python -m pip install --upgrade pywin32

# Verify COM
uv run python -c "import win32com.client; print('OK')"
```

- Ensure CAD is running
- Check CAD version (AutoCAD 2018+, ZWCAD 2020+)

### "Not connected"

Normal on startup - server auto-connects on first tool call.

If it fails:
1. Check CAD is running
2. Restart CAD
3. Use `manage_session` with `{"action": "connect"}`

## Operation Failures

### Workflow

### Before Committing
1. `uv run pytest tests/ -v` - All 181 tests must pass
2. `uv run ruff format src/` - Format
3. `uv run ruff check src/` - Lint
4. `uv run mypy src/` - Type check (must be clean)

### Drawing not visible

```python
zoom_extents()     # Fit view
list_layers()      # Check layer visibility
```

### Invalid coordinates

**Valid**:
```
"0,0"  "100.5,50.25"  "-10,-20"  "0,0,0"
```

**Invalid**:
```
"0, 0"   # Space
"(0,0)"  # Parentheses
"a,b"    # Non-numeric
```

## Configuration Issues

### Changes not taking effect

- Restart server after editing `src/config.json`
- Verify JSON syntax: `python -m json.tool src/config.json`

## Debug Tools

### Enable debug logging

Edit `src/config.json`:
```json
{"logging_level": "DEBUG"}
```

View logs:
```powershell
# Setup
uv sync --dev
uv run python -m pip install --upgrade pywin32

# Run
uv run python src/server.py

# Test
uv run pytest tests/ -v                      # 181 tests
npx -y @modelcontextprotocol/inspector uv run python src/server.py

# Quality
uv run ruff format src/ && uv run ruff check src/ && uv run mypy src/
```

Browse to `http://localhost:3000` to test tools interactively.

### Test connection directly

```python
from adapters import AutoCADAdapter
adapter = AutoCADAdapter("autocad")
adapter.connect()
print(f"Connected: {adapter.is_connected()}")
```

## Common Fixes

### ModuleNotFoundError

```powershell
uv sync --dev
uv run python -m pip install --upgrade pywin32
```

### Execution policy error

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
