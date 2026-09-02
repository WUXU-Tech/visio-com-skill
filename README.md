# Visio COM Skill

A minimal clean-room skill for generating editable Microsoft Visio `.vsdx` diagrams with local Windows COM automation.

This repository contains only a PowerShell renderer and skill instructions. It does not include third-party diagram examples, generated Visio files, local logs, or machine-specific paths.

## Requirements

- Windows
- Microsoft Visio Desktop
- PowerShell 5.1 or newer

## Install

For Codex:

```powershell
Copy-Item -Recurse . "$env:USERPROFILE\.codex\skills\visio-com-skill"
```

For Claude Code:

```powershell
Copy-Item -Recurse . "$env:USERPROFILE\.claude\skills\visio-com-skill"
```

## Use

Create a JSON spec, then run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\render_visio_diagram.ps1 -SpecPath .\diagram.json -OutputPath .\diagram.vsdx
```

The renderer opens Visio visibly by default and leaves the document open for review. For a background test run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\render_visio_diagram.ps1 -SpecPath .\diagram.json -OutputPath .\diagram.vsdx -Hidden -CloseWhenDone
```

## Minimal JSON

```json
{
  "page": { "width": 11, "height": 8.5 },
  "items": [
    { "type": "box", "x1": 1, "y1": 5.8, "x2": 3, "y2": 6.6, "text": "Input" },
    { "type": "box", "x1": 5, "y1": 5.8, "x2": 7, "y2": 6.6, "text": "Output" },
    { "type": "line", "x1": 3, "y1": 6.2, "x2": 5, "y2": 6.2, "arrow": true }
  ]
}
```

## Supported Items

- `text`
- `box`
- `diamond`
- `ellipse`
- `line`
- `connector`
- `frame`
- `foldBars`

## License

MIT
