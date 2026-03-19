# WMDraw17 — Technical Drawing Engine

A VB.NET library for generating vectorial engineering and technical drawings, with support for multiple output formats and COM interop for Excel/Office integration.

## Overview

WMDraw17 provides a flexible drawing engine built around a collection of `Drawable` objects rendered into a `Drawing` context. Coordinates can be specified in four interchangeable reference systems, and output can target PNG files, clipboard, DXF (AutoCAD), XAML, or a WPF Canvas.

The library is primarily used via its COM wrapper (`WMDraw17COM`), which exposes the drawing engine to Excel VBA and other COM-compatible hosts.

## Solution Structure

| Project | Type | Role |
|---|---|---|
| **WMDraw17** | Class Library (.NET 4.5) | Core drawing engine |
| **WMDraw17COM** | COM Library | COM/Excel interop wrapper |
| **TestWMDraw17_01** | WPF Application | Interactive test harness |
| **WMCore** | External Class Library | Shared utility dependency |
| **WMDrawCom2** | COM Library | Experimental alternative COM wrapper |

## Core Architecture

### Drawing Engine (`Drawing.vb`)

The `Drawing` class is the main orchestrator. It:
- Holds a collection of `Drawable` objects
- Manages coordinate transformation between reference systems
- Renders to a `ContextObject` (canvas, file path, etc.)
- Handles auto-scaling, margin management, and 300 DPI output

### Coordinate Reference Systems

Coordinates can be given in any of four systems, which are automatically converted at render time:

| Reference | Description |
|---|---|
| `world` | Engineering/model space (e.g. meters, mm of the structure) |
| `contextMillimeters` | Physical millimeters on the drawing canvas |
| `contextUnits` | Pixels of the output context |
| `contextFraction` | Relative position 0–1 across the canvas |

### Drawables (`Drawables/`)

All drawing elements implement the `Drawable` interface, which provides:
- Delegate functions for coordinate/size conversion (injected by the parent `Drawing`)
- `pen` and `fill` properties
- `zIndex` for rendering order
- `boundingRectangle()` for auto-scaling
- `draw()` to render into the current context

**Geometric primitives:**
`Point`, `Line`, `Rectangle`, `Polygon`, `Ellipse`

**Structural/engineering elements:**
`Beam`, `BeamUniformLoad`, `ForceArrow`, `MomentArc`, `Support`, `Screw`

**Annotations:**
`Text`, `DimensionLine`, `DimensionAngular`, `Gridline`, `ScaleBar`

**Decorative:**
`WoodGrainSymbol`, `Arrow` (configurable tip styles, force magnitude display)

### Geometry (`Geometry.vb`, namespace `WMDraw.Geom`)

Low-level geometric types: `Point`, `Vector`, `rectangle`, `polygon`, `line` — with operations such as dot product, cross product, rotation, and magnitude.

### Helpers (`WMHelpers.vb`)

Utility module providing:
- `SetSigFigs()` / `SetSigFigsString()` — round/format to significant figures
- `rainbowColor()` / `HSLtoRGB()` — color scale generation

### COM Wrapper (`WMDraw17COM/WMDraw.vb`)

`c_WMComDraw` wraps the core `Drawing` object for COM clients. Exposes methods such as `addLine()`, `addText()`, `addSupport()`, `drawToFile()`, and `clear()`. Enums are re-declared for COM visibility. The assembly is strong-named and registered via RegAsm in a post-build step.

## Output Formats

| Format | Description |
|---|---|
| PNG file | High-quality raster output (300 DPI) |
| PNG clipboard | Copy to clipboard for pasting into Excel |
| DXF file | AutoCAD-compatible vector format |
| XAML | WPF markup output |
| WPF Canvas | Live rendering into a WPF UI element |

## Dependencies

- **.NET Framework 4.5**
- **WPF** — `PresentationCore`, `PresentationFramework`, `WindowsBase`
- **GDI+** — `System.Drawing` (color and imaging)
- **WMCore** — internal shared utility library (`..\..\vb.net_productive\WMCore\`)
- Strong-named assemblies (`.snk`) for COM registration

## Getting Started

### Prerequisites

- Visual Studio 2019 or later
- .NET Framework 4.5 SDK
- WMCore project checked out at `..\..\vb.net_productive\WMCore\` relative to this repo

### Build

Open `WMDraw17.sln` in Visual Studio and build the solution. The COM projects (`WMDraw17COM`) require running Visual Studio as Administrator to register the assemblies via RegAsm during the post-build step.

### Using from VBA / Excel

After building and registering `WMDraw17COM.dll`, reference it in your VBA project via Tools > References. Use `c_WMComDraw` to create and render drawings:

```vb
Dim draw As New c_WMComDraw
draw.addLine 0, 0, 100, 100
draw.drawToFile "C:\output.png"
```

### Using from VB.NET / C\#

Reference `WMDraw17.dll` directly. Create a `Drawing`, add `Drawable` objects, set a `ContextObject`, and call `draw()`:

```vb
Dim d As New Drawing()
d.add(New Drawables.Line(New Point(0, 0), New Point(100, 50)))
d.drawToContext(ctx)
```

## Test Project

`TestWMDraw17_01` is a WPF application that exercises the full API — points, lines, beams, ellipses, force arrows, supports, dimensions, and coordinate reference mixing. The project also includes `PasteImage.xlsm` for testing clipboard/PNG export into Excel.

## Project History (recent)

- **Dowel-Grid** — added dowel grid drawing capability
- **300 DPI support** — high-resolution PNG export; tested with fraction, unit, mm, and world coordinates in Excel
- **PNG file output** — PNG export to file and clipboard insertion into Excel
- **WMCore integration** — updated dependency on shared utility library
- **Code reorganisation** — `Drawable.vb` moved to `Drawables/` folder; `Arrow.vb` extracted to its own file; `WMHelpers.vb` added
- **DoveTail UI** — new `Frm_DoveTail` Windows Forms component added to COM project
