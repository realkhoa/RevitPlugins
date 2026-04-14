# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

This is a collection of **Revit 2026 API add-ins** (DIM = Dimension tools). Each module is an independent .NET 8 project that registers a button on a shared "DIM" ribbon tab.

## Build

```bash
cd <ModuleName> && dotnet build
```

Each module compiles to its own DLL. No central solution or CI/CD.

## Project Structure

```
<ModuleName>/
  Class1.cs      # App (ribbon) + Command (IExternalCommand)
  <ModuleName>.csproj

Key modules:
  AlignedDIM/           Live multi-pick aligned dimension
  AlignedDIMParam/      AlignedDIM + prompt for FamilyLabel
  AngularDIM/           Angular dimension via edge intersection
  AngularDIMParam/      AngularDIM + prompt for FamilyLabel
  QParam/               Pick existing DIM, assign FamilyLabel parameter
  MultiAlignedDIM/      Auto-DIM pre-selected elements by direction grouping
```

## Ribbon Pattern (all modules)

`App : IExternalApplication` creates a shared "DIM" tab with a "Dimension" panel. Each module's button uses `Assembly.GetExecutingAssembly().Location` as the path.

**Important**: The `PushButtonData` constructor's 4th argument must exactly match `{Namespace}.{ClassName}` of the class implementing `IExternalCommand`.

## Revit 2026 API Notes

**Namespace conflicts** — always disambiguate:
- `TaskDialog` ambiguous with `System.Windows.Forms.TaskDialog` → use `using TaskDialog = Autodesk.Revit.UI.TaskDialog;`
- `View` ambiguous with `System.Windows.Forms.View` → use `using View = Autodesk.Revit.DB.View;`
- WPF types (`System.Windows.Window`, `System.Windows.Controls.*`) not available — use WinForms or fully qualify
- WinForms types in csproj require `<UseWindowsForms>true</UseWindowsForms>`

**SpecTypeId / GroupTypeId** — these are `abstract sealed` static classes. Do NOT declare variables of these types. Use `ForgeTypeId`:
```csharp
ForgeTypeId targetSpecType = isAngular ? SpecTypeId.Angle : SpecTypeId.Length;
ForgeTypeId groupType = isAngular ? GroupTypeId.Text : GroupTypeId.Geometry;
doc.FamilyManager.AddParameter(name, groupType, targetSpecType, false);
```

**Dimension creation** differs between project and family:
```csharp
// Project document
m_doc.Create.NewDimension(view, dimLine, refArray);
m_doc.Create.NewDimension(view, line, refArray);

// Family document
doc.FamilyCreate.NewDimension(view, dimLine, refArray);
doc.FamilyCreate.NewAngularDimension(view, arc, r1, r2);  // Angular

// Project angular dimension (Revit 2022+)
AngularDimension.Create(doc, view, arc, refs, dimensionType);
```

**DimensionType** — use `dim.DimensionType` (already `DimensionType`, not `ElementId`):
```csharp
if (dim.DimensionType is DimensionType dt && dt.StyleType == DimensionStyleType.Angular)
```

**Edge references** — not always available on all edge geometry. Fallback to face references via `PlanarFace.Reference`.

**Angular vs Linear** detection: check `DimensionStyleType.Angular` or `dim.Curve is Arc`.

**DIM only works in 2D views** — `doc.ActiveView.ViewType == ViewType.ThreeD` should be rejected.

## Adding a New Module

1. Copy an existing module folder and rename
2. Update namespace in `.csproj` and `Class1.cs`
3. Register the class in `PushButtonData` constructor (4th arg)
4. Add `<UseWindowsForms>true</UseWindowsForms>` if using WinForms dialogs


Always use Context7 when I need library/API documentation, code generation, setup or configuration steps without me having to explicitly ask.
