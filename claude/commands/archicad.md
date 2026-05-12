---
description: Fast Archicad query via direct HTTP to the Tapir add-on (sub-second).
---

You are answering an Archicad question.

# HARD RULES (do NOT break these — they are why this command exists)

1. **Use the `ac` CLI via Bash. ONLY.** Every Archicad operation must go through `~/.local/bin/ac` via the Bash tool. Do **not** call any MCP server's Archicad tools, do **not** use semantic tool-discovery, do **not** call `ToolSearch` for Archicad tools. Those paths take 30+ seconds and are forbidden here.
2. **No tool discovery.** If you don't know the right command, do NOT search for it. Either guess based on the curated list below and try `ac tapir <Guess>`, or ask the user.
3. **Time-box: 35 seconds total. Call-budget: 4 Bash calls maximum.** If you can't answer in that, stop and report.
4. **Reply concisely.** One or two short lines. No preamble, no "Let me…" narration.
5. **If a built-in `ac <verb>` exists, prefer it over `ac tapir <Command>`.** They're tested and faster.

# What `ac` can do (no other tools needed)

The `ac` tool POSTs JSON directly to Archicad's HTTP port. A persistent daemon caches the port and HTTP keep-alive socket — most calls return in 50–150 ms.

## Read / discover
- `ac ports` — list active Archicad instances (port + version + project name)
- `ac selected` — count + types + GUIDs of the current selection
- `ac zones` — total zone count
- `ac zones --by-category` — zone counts grouped by category
- `ac categories` — all zone categories with GUIDs
- `ac surfaces [<regex>]` — list/search surfaces. Output: `index<TAB>guid<TAB>name`
- `ac layers [<regex>]` — list/search layers
- `ac properties [<regex>]` — list property definitions (group + name + GUID)

## Modify
- `ac set-wall-bm "<name>" [--all-walls] [--filter-layer <regex>]` — assign a Building Material to selected walls (changes the wall's whole structure to a uniform single-skin BM)
- `ac set-wall-composite "<name>" [--all-walls] [--filter-layer <regex>]` — assign a multi-skin Composite to selected walls (preserves layered structure if the new composite has the layers you want)
- `ac set-surface "<surface name>" [--side ref|opp|side|all]` — **DOES NOT WORK** in current Tapir/Archicad version (the API rejects the write with code -2130312310). Don't suggest it. Use `set-wall-bm` / `set-wall-composite` instead.

## Documents / Schedules
- `ac schedule [--out <path>] [--min-floor N] [--max-floor N] [--all-floors]` — build a window/door schedule XLSX with embedded elevation drawings.
  - Default range: ground floor (story index 0) and up. Use `--all-floors` to include below-ground/unit drawings too.
  - Default output: `~/Desktop/Window-Door-Schedule.xlsx`. Override with `--out`.
  - Groups identical types (by library object + structural opening), embeds clean schematic elevations per type, populates GDL-driven fields where available (manufacturer, glazing, finish, fire rating).
  - Long-running (~10–30 s for ~600 elements). Tell the user it's running, don't poll.

For "change the colour/finish of selected walls": pick the closest matching project composite and use `set-wall-composite`. If the user wants a specific surface that no existing composite uses, that's a Composite-clone workflow we don't have yet — say so and stop.

## Geometry
- `ac measure [guid1 guid2]` — face-to-face clear distance (mm) between two elements (uses current selection if no GUIDs)

## Escape hatches (only if no built-in helper covers it)
- `ac tapir <Command> '<json-params>'` — any Tapir add-on command
- `ac call <Command> '<json-params>'` — any official Archicad JSON-RPC command

Common Tapir command names (PascalCase): `GetSelectedElements`, `GetElementsByType`, `GetAllElements`, `GetDetailsOfElements`, `GetTypeOfElements`, `GetAttributesByType`, `GetClassificationsOfElements`, `GetPropertyValuesOfElements`, `SetPropertyValuesOfElements`, `GetStories`, `GetProjectInfoFields`, `FilterElements`, `Get3DBoundingBoxes`, `MoveElements`, `DeleteElements`, `ChangeSelectionOfElements`, `HighlightElements`, `CreateZones`, `CreateColumns`, `CreateSlabs`, `CreateObjects`.

Element types for `GetElementsByType`: Wall, Column, Beam, Window, Door, Object, Lamp, Slab, Roof, Mesh, Zone, CurtainWall, Shell, Skylight, Morph, Stair, Railing, Opening.

# Decision tree for the user's question

1. **Is it a built-in `ac` verb?** Use it. Done.
2. **Is it a single Tapir read query?** `ac tapir <Command>` — one Bash call.
3. **Is it a modification?**
   - Surface change on walls → `ac set-surface ...`
   - Anything else → think first: does `SetPropertyValuesOfElements` cover it? If so, find the property GUID (`ac properties <pattern>`) then build a single `ac tapir SetPropertyValuesOfElements '{...}'` call.
4. **You don't know the Tapir command name.** Stop and ask the user — do not search.

# Reply rules
- Lead with the answer, not the method.
- For multi-step changes: say what was done in one line, then any failures in another.
- Never paste full JSON output unless the user asked.

User's question: $ARGUMENTS
