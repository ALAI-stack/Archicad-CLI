# Archicad CLI (`ac`)

A fast, local command-line tool that lets Claude Code (and you) talk directly to your running Archicad model. It bypasses the MCP tool-discovery layer entirely most queries return in **under a few seconds (dependant on task complexity)**, and the `/archicad` slash command in Claude Code gives a curated, low-token workflow for BIM tasks.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Status](https://img.shields.io/badge/status-alpha-orange.svg)]()
[![Platform](https://img.shields.io/badge/platform-macOS-lightgrey.svg)]()

> **Disclaimer:** This project is in early-stage development and is actively being updated. APIs, commands, and the schedule layout may change. Please use with caution and back up your model before running any write operations.

## Key Features

- **Sub-second BIM queries.** `ac` POSTs JSON directly to Archicad's HTTP port (exposed by the Tapir Archicad Add-On). No MCP, no semantic tool discovery, no schema fetching — just one Bash call per query.
- **Token-efficient by design.** A curated `/archicad` slash command bundles the entire toolset inline (~1.5 k tokens, paid once per session) instead of letting Claude dynamically load tool schemas. A 5-query session is typically **3–5× cheaper** than an MCP-based equivalent.
- **Auto-spawned background daemon.** First call spawns a tiny Python daemon that caches the Archicad port + holds an HTTP keep-alive socket. Subsequent calls return in ~10–30 ms. Daemon idles out after 10 minutes.
- **Built-in helpers for common tasks.** `ac zones`, `ac selected`, `ac surfaces`, `ac measure`, `ac set-wall-bm`, `ac set-wall-composite`, plus generic `ac tapir <Command>` for anything else.
- **Window/Door schedule generator.** `ac schedule` produces an XLSX with embedded schematic elevation drawings per type, dimensions, fixed-light tags, swing markers, and a tabular data section. Reads layout from `ifc_optypestr` and `gs_optype_*` GDL params; data from `gs_list_*` fields populated by Archicad's Window/Door Settings.
- **`ac doctor` self-check.** Walks through every prerequisite (CLI installed, slash command installed, Archicad reachable, daemon health) and prints a clear pass/fail summary.
- **100% local & private.** All communication happens between your Mac and Archicad on localhost. No cloud calls, no API keys, no data leaves your machine.

## Installation & Setup

### 1. Prerequisites

- **macOS** (Linux/Windows support not yet tested).
- **Claude Code** installed.
- **Archicad with the [Tapir Archicad Add-On](https://github.com/ENZYME-APD/tapir-archicad-automation) installed.** The Add-On is what opens the HTTP port (19723–19743) that the CLI talks to. Without it, `ac` has nothing to connect to.
- **`uv`** (modern Python package manager). The install script will install it for you if missing.

### 2. Install

```bash
git clone https://github.com/ALAI-stack/Archicad-CLI.git ~/Archicad-CLI \
  && cd ~/Archicad-CLI \
  && bash install.sh
```

The installer:
1. Installs `uv` if missing.
2. Copies the `ac` CLI to `~/.local/bin/ac`.
3. Merges fast-path rules into your global `~/.claude/CLAUDE.md` (between markers, preserving anything else you have there).
4. Installs the `/archicad` slash command to `~/.claude/commands/archicad.md`.
5. Installs the schedule helper script to `~/.local/share/ac-scripts/`.

### 3. Verify

Open a project in Archicad, then:

```bash
ac doctor
```

This self-check reports on every prerequisite. If anything is missing, it tells you what.

Then restart Claude Code and try:

```
/archicad how many zones do I have?
```

## Usage

### From the terminal

```bash
ac ports                              # list active Archicad instances
ac selected                           # current selection summary (count + types + GUIDs)
ac zones                              # total zone count
ac zones --by-category                # zone counts grouped by Zone Category
ac surfaces [<regex>]                 # list/search surfaces
ac layers [<regex>]                   # list/search layers
ac measure                            # face-to-face distance between 2 selected elements
ac set-wall-composite "<name>"        # assign a Composite to selected walls
ac schedule                           # build window/door schedule → ~/Desktop/Window-Door-Schedule.xlsx
ac tapir <Command> '<json-params>'    # any Tapir add-on command (escape hatch)
ac call  <Command> '<json-params>'    # any official Archicad command (escape hatch)
ac --help                             # full command list
```

### From Claude Code

The `/archicad` slash command makes Claude use `ac` automatically:

```
/archicad how many doors are on the first floor
/archicad list all surfaces matching "paint"
/archicad change selected walls to composite IW-04
/archicad build a window/door schedule for ground floor and up
```

## How It Works

```
Claude Code  ──▶  /archicad slash command  ──▶  Bash: ac <verb>
                                                       │
                                                       ▼
                                              ~/.local/bin/ac
                                                       │
                                          (auto-spawns daemon on first call)
                                                       │
                                                       ▼
                                                  /tmp/ac.sock  ◀──  cached port,
                                                       │              persistent HTTP socket
                                                       ▼
                                       HTTP POST → localhost:19724
                                                       │
                                                       ▼
                                               Tapir Archicad Add-On
                                                       │
                                                       ▼
                                                    Archicad
```

- **`ac` CLI** — small Python script (stdlib-only). Parses arguments, formats the JSON-RPC request, talks to the daemon (or directly to Archicad as a fallback).
- **Daemon** — same script invoked as `ac --daemon`. Listens on a Unix socket at `/tmp/ac.sock`. Holds the discovered Archicad port and a persistent HTTP connection.
- **`/archicad` slash command** — instructs Claude to use `ac` via Bash, with a curated command list inline so Claude never needs to call `ToolSearch` or `archicad_discover_tools`.
- **`CLAUDE.md` rules** — global Claude rules scoped between `<!-- ARCHICAD-FAST-START -->` / `<!-- ARCHICAD-FAST-END -->` markers so they merge cleanly with anything else you already have.

## Updating / Uninstalling

```bash
cd ~/Archicad-CLI && bash update.sh      # git pull + reinstall
cd ~/Archicad-CLI && bash uninstall.sh   # clean removal
```

## Contributing

Contributions are welcome! This project is in active development and there are plenty of rough edges. Useful directions to help:

- Cross-platform support (Linux / Windows).
- More built-in `ac` verbs for common architectural workflows.
- Custom Tapir command support in non-default namespaces.

Open an issue to discuss a change, or send a PR.

## License

MIT — see [LICENSE](./LICENSE). Anyone can use, modify, and redistribute this code, including commercially; just keep the LICENSE file attached.

## Acknowledgements

This project was inspired by the [Tapir Archicad MCP](https://github.com/SzamosiMate/tapir-archicad-MCP) by Máté Szamosi, which originated the idea of bridging AI assistants to Archicad. This CLI is an independent reimplementation that talks directly to the Tapir Archicad Add-On over HTTP, skipping the MCP layer for speed.

It depends on the [Tapir Archicad Add-On](https://github.com/ENZYME-APD/tapir-archicad-automation) by ENZYME-APD for the JSON-RPC API exposed by Archicad.
