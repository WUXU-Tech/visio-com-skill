---
name: visio-com-skill
description: Generate editable Microsoft Visio .vsdx diagrams on Windows through local Visio COM automation. Use when an agent needs to create or revise Visio workflow diagrams, flowcharts, architecture diagrams, routed connectors, legends, or precise diagram reconstructions from a JSON specification.
---

# Visio COM Skill

Create editable `.vsdx` files using Microsoft Visio Desktop through local PowerShell COM automation.

## Requirements

- Windows
- Microsoft Visio Desktop installed and COM-registered
- PowerShell 5.1 or newer

Check Visio COM availability:

```powershell
[type]::GetTypeFromProgID("Visio.Application")
```

If this returns `$null`, Visio COM is unavailable. Offer another editable format such as SVG, Mermaid, or draw.io.

## Workflow

1. Convert the user's request into a JSON diagram specification.
2. Save the JSON specification in the active workspace.
3. Run:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "<skill_dir>\scripts\render_visio_diagram.ps1" -SpecPath "<diagram.json>" -OutputPath "<diagram.vsdx>"
   ```

4. By default, the renderer opens Visio visibly and leaves the document open for review.
5. For automated checks, use `-Hidden -CloseWhenDone`.

## JSON Format

Use an `items` array for precise diagram control. Coordinates are in Visio inches.

Supported item types:

- `text`
- `box`
- `diamond`
- `ellipse`
- `line`
- `connector`
- `frame`
- `foldBars`

Minimal specification:

```json
{
  "fileName": "diagram.vsdx",
  "pageName": "Diagram",
  "page": { "width": 11, "height": 8.5 },
  "items": [
    { "type": "text", "x1": 3.5, "y1": 7.6, "x2": 7.5, "y2": 8.1, "text": "Workflow", "size": 18, "bold": true },
    { "type": "box", "x1": 1, "y1": 5.8, "x2": 3, "y2": 6.6, "text": "Input", "fill": "#F5F5F5" },
    { "type": "diamond", "cx": 5.1, "cy": 6.2, "w": 1.4, "h": 0.9, "text": "Model", "fill": "#FFF6D7" },
    { "type": "box", "x1": 7, "y1": 5.8, "x2": 9, "y2": 6.6, "text": "Output" },
    { "type": "line", "x1": 3, "y1": 6.2, "x2": 4.4, "y2": 6.2, "arrow": true },
    { "type": "line", "x1": 5.8, "y1": 6.2, "x2": 7, "y2": 6.2, "arrow": true }
  ]
}
```

## Layout Notes

- Match the page ratio to the source diagram before placing shapes.
- Keep long feedback lines as orthogonal `connector` segments.
- Separate feedback lines from section divider lines.
- Use `frame` with `"dashed": true` for workflow boundaries.
- Use `diamond` for decision/model nodes instead of rotated rectangles.
- Save `.vsdx` first. Avoid preview export unless the user specifically asks for it.
