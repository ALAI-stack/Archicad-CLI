<!-- ARCHICAD-FAST-START -->
# Archicad fast-path rules

For any question about Archicad, BIM models, zones, walls, doors, slabs, stories, or anything in a `.pln` file:

**Use the `ac` CLI via Bash.**

The `ac` tool is at `~/.local/bin/ac`. It POSTs JSON directly to Archicad's HTTP port (sub-second responses). A persistent background daemon auto-spawns on first call so subsequent calls reuse a cached port and an open HTTP keep-alive connection (~10–30 ms each).

## Built-in commands
- `ac ports` — list active Archicads (port + version + project name)
- `ac zones` — total zone count
- `ac zones --by-category` — zones grouped by Zone Category (Standard Room, etc.)
- `ac categories` — list all Zone Category attributes
- `ac measure [guid1 guid2]` — face-to-face clear distance between two elements (uses current selection if no GUIDs given)
- `ac selected` — summary of current selection (port + project + types + GUIDs)
- `ac surfaces [<regex>]` / `ac layers [<regex>]` / `ac properties [<regex>]` — list/search attributes
- `ac set-wall-bm "<name>"` / `ac set-wall-composite "<name>"` — assign BM/Composite to selected walls
- `ac schedule [--out <path>]` — build window/door schedule XLSX (ground-floor and up by default)
- `ac tapir <Command> '<json-params>'` — any Tapir add-on command
- `ac call <Command> '<json-params>'` — any official Archicad command
- `ac status` — show daemon + port-cache state
- `ac stop` — stop the background daemon (auto-respawns next call)

Use PascalCase command names with `ac tapir` (e.g. `GetElementsByType`, `GetStories`, `GetProjectInfoFields`).

## Speed rules
1. Default to `ac` via Bash.
2. Hard time-box: 35 seconds per Archicad question unless the user says otherwise.
3. Hard call-budget: max ~4 Bash calls per question.
4. If the user gives a tighter time-box (e.g. "10 seconds"), respect it and stop with a status.
5. Reply concisely — one or two short lines. No preamble.

## Common Tapir commands
GetElementsByType, GetAllElements, GetSelectedElements, GetDetailsOfElements, GetZoneBoundaries, GetClassificationsOfElements, GetStories, GetProjectInfoFields, GetAttributesByType, GetPropertyValuesOfElements, GetHotlinks, GetGeoLocation, FilterElements, Get3DBoundingBoxes, MoveElements, DeleteElements, ChangeSelectionOfElements, HighlightElements, CreateZones, CreateColumns, CreateSlabs, CreateObjects.

## Element types for GetElementsByType
Wall, Column, Beam, Window, Door, Object, Lamp, Slab, Roof, Mesh, Zone, CurtainWall, Shell, Skylight, Morph, Stair, Railing, Opening.
<!-- ARCHICAD-FAST-END -->
