# 05 - Tool Reference

## Summary

**7 unified MCP tools** providing access to **55 CAD commands** via a shorthand dispatch format:

| Unified Tool | Actions | Category |
|--------------|---------|----------|
| `manage_session` | connect, disconnect, status, zoom_extents, undo, redo, screenshot, export_view, check_running, open_dashboard, list_supported | Connection & Control |
| `draw_entities` | line, circle, arc, rect, polyline, spline, text, dimension, leader, mleader | Drawing |
| `manage_blocks` | list, info, insert, create, get_attrs, set_attrs | Blocks |
| `manage_layers` | create, delete, rename, on, off, set_color, is_on, list, info | Layers |
| `manage_files` | save, new, close, list, switch | Files |
| `manage_entities` | select, move, rotate, scale, set_color, set_layer, set_color_bylayer, copy, paste, delete | Entities |
| `export_data` | scope=all/selected, format=json/excel | Export |

---

## Shorthand Format

All tools accept operations as a **plain-text shorthand** (one per line), which is far more token-efficient than JSON:

```
# draw_entities
line|0,0|100,0|red|walls
circle|50,40|10|blue

# manage_layers
create|walls|red|50
off|Defpoints,notes

# manage_entities
select|layer|walls
move|A1B2,C3D4|10|5
```

JSON arrays are also accepted for backwards compatibility.

---

## manage_session

Connection lifecycle, view control, and history.

```
connect                          # auto-detect and connect
disconnect                       # release COM connection
status                           # current connection status
check_running                    # detect CAD without launching
list_supported                   # list available CAD types
zoom_extents                     # fit view to all entities
undo                             # undo 1 action
undo|{"count": 3}                # undo 3 actions (JSON format)
redo                             # redo 1 action
screenshot                       # capture window (includes UI chrome)
export_view                      # render drawing internally (works obscured)
open_dashboard                   # open web dashboard in browser
```

---

## draw_entities

Create geometric entities. Shorthand: `type|param1|param2|...|color|layer`

```
line|start|end|color|layer                    → line|0,0|10,10|red|walls
circle|center|radius|color                    → circle|5,5|3|blue
arc|center|radius|start_angle|end_angle       → arc|0,0|5|0|90
rect|corner1|corner2|color                    → rect|0,0|20,15
polyline|pts(;sep)|closed|color               → polyline|0,0;10,10;20,0|closed
spline|pts(;sep)|closed|color                 → spline|0,0;5,10;10,0
text|pos|text|height|color                    → text|5,5|Hello World|2.5
dimension|start|end|color                     → dimension|0,0|10,0
leader|pts(;sep)|text|height|color|layer      → leader|0,0;10,10|My note|2.5|red
leader|group1~~group2|text|...                → leader|0,0;10,10~~20,0;10,10|Label
```

**DEFAULTS:** `color=white`, `layer=0`

**Leaders:** first point = arrowhead, last point = text attach. Use `~~` for multi-arrow.

---

## manage_layers

```
create|name|color|lineweight     → create|walls|red|50
delete|name                      → delete|temp
rename|old|new                   → rename|Layer1|furniture
on|names(,sep)                   → on|walls,doors
off|names(,sep)                  → off|Defpoints
set_color|name|color             → set_color|0|white
is_on|name                       → is_on|walls
list                             # list all layer names
info                             # full layer details (color, locked, frozen)
```

**DEFAULTS:** `color=white`, `lineweight=25`

---

## manage_blocks

```
list                                           # list all block definitions
info|block_name|include                        → info|Door|both
insert|name|point|scale|rotation|layer|color  → insert|Door|10,20|1.5|90|walls|red
create|name|handles|point|description         → create|MyBlock|A1,B2|0,0|Desc
get_attrs|handle                              → get_attrs|A1B2C3
set_attrs|handle|{"TAG": "value"}             → set_attrs|A1B2C3|{"POLOS": "4P"}
```

**`include`:** `info` (default) | `references` | `both`

---

## manage_entities

```
select|by|value               → select|layer|walls
move|handles|dx|dy            → move|A1,B2|10|5
rotate|handles|angle|cx|cy   → rotate|A1|45|0|0
scale|handles|factor|cx|cy   → scale|A1|2.0|0|0
set_color|handles|color      → set_color|A1,B2|red
set_layer|handles|layer      → set_layer|A1|walls
set_color_bylayer|handles    → set_color_bylayer|A1,B2
copy|handles                 → copy|A1,B2
paste|base_point             → paste|100,200
delete|handles               → delete|A1,B2
```

**`by`:** `color` | `layer` | `type`
**`handles`:** comma-separated entity handles (e.g. `A1B2,C3D4`)

---

## manage_files

```
save|path_or_filename|format  → save|/path/file.dwg
save|filename                 → save|backup.dwg
new                           # create new drawing
close|save_changes            → close|true
list                          # list open drawings
switch|drawing_name           → switch|floor_plan.dwg
```

**`format`:** `dwg` (default) | `dxf` | `pdf`

---

## export_data

Extract entity data as JSON or export to Excel (3 sheets: Entities, Layers, Blocks).

```
scope=all, format=json              # all entities as JSON
scope=all, format=excel             # export to [drawing_name]_data.xlsx
scope=selected, format=json         # selected entities as JSON
scope=selected, format=excel, filename=selection.xlsx
```

**Excel columns:** `Handle`, `ObjectType`, `Layer`, `Color`, `Length`, `Area`, `Radius`, `Circumference`, `Name`

---

## Common Parameters

### Coordinates
```
"0,0"        # 2D
"0,0,0"      # 3D
"100.5,-20"  # Floats and negatives
```

### Colors
`red`, `blue`, `green`, `yellow`, `cyan`, `magenta`, `white`, `black`, `gray`, `orange`

### CAD Types
`autocad`, `zwcad`, `gcad`, `bricscad`

### Entity Types (for select|type|...)
`LINE`, `CIRCLE`, `ARC`, `LWPOLYLINE`, `POLYLINE`, `TEXT`, `MTEXT`, `INSERT`, `DIMENSION`, `SPLINE`, `POINT`, `HATCH`
