# Web Dashboard

## Overview

The multiCAD-mcp server includes an integrated web dashboard for real-time monitoring and control of CAD applications. The dashboard provides a user-friendly interface for viewing CAD status, managing layers and blocks, and monitoring entities.

## Access

When the server is running, the dashboard is available at (default from config.json: 6666):

```
http://localhost:6666
```

The dashboard provides a real-time monitor of the CAD state. You can manually refresh the data using the "Refresh Now" button.

## Features

### 1. Status Cards

**CAD Connection Status**
- Shows current CAD application type (AutoCAD, ZWCAD, GstarCAD, BricsCAD)
- Displays connection status (Connected/Disconnected)
- Indicates active drawing name

**Statistics**
- Total entity count in active drawing
- Active layer information
- Block library size

### 2. Layers Section

**Layer Management**
- View all layers in the active drawing
- Toggle layer visibility directly from dashboard
- Color-coded layer information
- See locked/unlocked status

**Real-time Updates**
- Automatically syncs with CAD application
- Reflects changes made via API or CAD UI

### 3. Blocks Section

**Block Library**
- Browse all blocks in active drawing
- View block properties
- See block references count
- Copy block information

### 4. Entities Section

**Entity Browser**
- List all entities in active drawing
- Filter by entity type (line, circle, arc, text, etc.)
- View entity properties
- Entity count statistics

## API Endpoints

The dashboard uses these backend API endpoints:

### GET `/api/cad/status`

Get current CAD connection status.

```json
{
  "success": true,
  "status": {
    "connected": true,
    "cad_type": "autocad",
    "drawings": ["Drawing1.dwg"],
    "current_drawing": "Drawing1.dwg",
    "supported": ["autocad", "zwcad", "gcad", "bricscad"]
  }
}
```

### GET `/api/cad/layers`

Get layers from active drawing.

```json
{
  "success": true,
  "layers": [
    {
      "name": "0",
      "color": 7,
      "frozen": false,
      "locked": false
    }
  ]
}
```

### GET `/api/cad/blocks`

Get blocks from active drawing.

```json
{
  "success": true,
  "blocks": [
    {
      "name": "Block1",
      "count": 5
    }
  ]
}
```

### GET `/api/cad/entities`

Get entities from active drawing.

```json
{
  "success": true,
  "entities": [
    {
      "handle": "ABC",
      "type": "LINE",
      "layer": "0",
      "color": 256
    }
  ]
}
```

### POST `/api/cad/refresh`

Manually trigger a dashboard cache refresh.

```json
{
  "success": true,
  "detail": "Refresh completed"
}
```

### GET `/api/health`

Health check endpoint.

```json
{
  "status": "ok",
  "version": "0.2.0"
}
```

## Architecture

### Thread Safety

The dashboard runs in a separate thread from the MCP server to avoid blocking CAD operations. Communication between threads is handled through a thread-safe cache:

- **MCP Thread**: Owns COM objects, updates cache after operations
- **Dashboard Thread**: Reads from cache, provides UI

This design ensures CAD operations remain responsive even during heavy UI activity.

### Cache System

The `DashboardCache` class provides thread-safe read/write operations:

```python
cache = DashboardCache()
cache.update(connected=True, cad_type="autocad")  # From MCP thread
layers = cache.get("layers")  # From dashboard thread
```

Cache is automatically populated when:
- Server connects to CAD application
- Manual refresh is triggered
- MCP tools modify CAD data

## Configuration

The dashboard provides a real-time view of the CAD state. This is configurable in `src/config.json`:

- **Dashboard Port**: Change `dashboard.port` to your preferred port.
- **Manual Refresh**: Click the "Refresh Now" button to sync with current CAD state.

### Static Files

Dashboard static files are located in `src/web/static/`. The server automatically mounts these at `/static/`.

## Performance Considerations

### Large Drawings

For drawings with 10,000+ entities:

- Initial load may take 1-2 seconds
- Entity filtering is recommended
- Use manual refresh instead of auto-refresh for better responsiveness

### Network Optimization

- All data is served locally (no external requests)
- WebSocket support not implemented (uses polling via HTTP)
- Cache is regenerated on manual refresh

## Troubleshooting

### Dashboard Shows "No Connection"

1. Verify CAD application is running
2. Check server logs for connection errors
3. Try manual refresh

### Stale Data

1. Click "Refresh Now" button to force update
2. Verify CAD window is not minimized
3. Check for operation locks in CAD application

### Performance Issues

1. Click the "Refresh Now" button
2. Filter entity list by type
3. Verify CAD application performance

## Future Enhancements

Planned improvements for v0.3.0+:

- WebSocket support for real-time updates
- Entity property editor
- Layer/block creation from dashboard
- Drawing comparison tools
- Export drawing metadata
- Multi-drawing support
