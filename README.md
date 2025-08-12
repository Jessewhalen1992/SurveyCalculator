# SurveyCalculator

AutoCAD Map 3D plugin seed (.NET 8, x64).

## Build requirements
- Windows + Visual Studio 2022
- .NET 8 (net8.0-windows)
- AutoCAD 2025 (managed DLL references; Map API lives under the `\Map` subfolder)

Set your AutoCAD install path in **Directory.Build.props**:
```xml
<ACAD_DIR>C:\Program Files\Autodesk\AutoCAD 2025</ACAD_DIR>
```

## Build
Open src/SurveyCalculator/SurveyCalculator.csproj in VS and build Release | x64, or:

```css
msbuild .\src\SurveyCalculator\SurveyCalculator.csproj /p:Configuration=Release /p:Platform=x64
```

The DLL is copied to /deploy on post-build.

## Load & test in Map 3D
In AutoCAD Map 3D: NETLOAD → pick deploy\SurveyCalculator.dll

Use the commands below to perform the ALS workflow.

## Commands
- PLANCOGO: Build plan graph from numeric bearings/distances (manual entry or CSV); draws plan polyline and writes {DWG}_PlanGraph.json.
- EVILINK: Link evidence blocks → PlanID + Held/Floating (FDI auto‑Held, fdspike weaker); writes {DWG}_EvidenceLinks.json.
- ALSADJ: ALS‑style adjust; Held anchors fixed; traverse between anchors by plan numbers with compass closure; draws adjusted boundary + residuals; writes {DWG}_AdjustmentReport.csv.
- PLANLINK / PLANEXTRACT: Legacy commands kept for PDFIMPORT workflows.

## Usage (quick)
1) PLANCOGO → enter bearings/distances manually or load a CSV to create the plan boundary and {DWG}_PlanGraph.json.
2) EVILINK → window-select your evidence blocks; assign PlanIDs and Held/Floating; Save to {DWG}_EvidenceLinks.json.
3) ALSADJ → produces adjusted boundary on BOUNDARY_ADJ, residuals on QA_RESIDUALS, and CSV in your DWG folder.

Legacy PDFIMPORT + PLANLINK + PLANEXTRACT workflow is still available but not preferred.
