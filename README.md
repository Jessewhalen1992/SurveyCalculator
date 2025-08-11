# SurveyCalculator

AutoCAD Map 3D plugin seed (.NET Framework 4.8, x64).

## Build requirements
- Windows + Visual Studio 2022
- .NET Framework 4.8 targeting pack
- AutoCAD Map 3D (managed DLL references)

Set your Map 3D install path in **Directory.Build.props**:
```xml
<ACAD_DIR>C:\Program Files\Autodesk\AutoCAD Map 3D 2025</ACAD_DIR>
```

## Build
Open src/SurveyCalculator/SurveyCalculator.csproj in VS and build Release | x64, or:

```css
msbuild .\src\SurveyCalculator\SurveyCalculator.csproj /p:Configuration=Release
```

The DLL is copied to /deploy on post-build.

## Load & test in Map 3D
In AutoCAD Map 3D: NETLOAD → pick deploy\SurveyCalculator.dll

Run HELLOALS in the command line.

## Next steps
Replace the hello file with ALS commands (PLANLINK, PLANEXTRACT, EVILINK, ALSADJ).
