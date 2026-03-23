
![vbaXray logo](logo_vbaxray.png)
# vbaXray

A pure-VBA class that reads, inspects, and exports VBA source code directly from Office binary files — without opening them, without requiring any admin rights or registry changes.

---

## What it does

vbaXray parses the `vbaProject.bin` Compound Document Storage embedded in any macro-enabled Office file (`.xlsm`, `.xlam`, `.docm`, `.dotm`, `.pptm`, `.potm`, etc.) and returns:

- the project name and code page
- the name, type, and decompressed source of each module
- the ability to export all modules, organised by type

It works entirely in-process using `StgOpenStorage` / COM vtable dispatch. 

No manual unzipping is required - instead, it uses `Shell.Application` to create a copy of the source Office file, and then extract the contents. A cleaner approach would be to extract the vbaProject.bin to a byte array, but this was the quickest solution for now. I have written a pure-VBA ZIP extractor and will add it in a future update, but for now the polling approach is a workable fallback.

Alternatively, I would recommend using the excellent Excel-ZipTools (https://github.com/cristianbuse/Excel-ZipTools), by Cristian Buse, which provides a pure-VBA ZIP reader and can extract the `vbaProject.bin`.

---

## Quick start

1. Add `clsVBAXray.cls` to your VBA project.
2. Call `LoadFromFile` with the path to an Office file or a raw `vbaProject.bin`.
   
```vba
Dim xray As New clsVBAXray

' Load from an xlsm / docm / pptm  — or directly from a vbaProject.bin
xray.LoadFromFile "C:\Projects\MyMacroFile.xlsm"

' Inspect
Debug.Print "Project Name: " & xray.ProjectName        ' e.g. "VBAProject"
Debug.Print "Module Count: " & xray.ModuleCount        ' e.g. 5

' List every module
Dim i As Long
For i = 1 To xray.ModuleCount
    Debug.Print i, xray.ModuleName(i)
Next i

' Read the source of module 1
Debug.Print "Source of Module 1:" & vbCrLf & xray.SourceCode(1)

' Export everything to disk (organised into sub-folders by type)
xray.ExportAll "C:\Export\MyProject\"

' Or export a single module by name
xray.ExportModuleByName "Module1", "C:\Export\MyProject\"
```
---

## How it works

Office macro-enabled files (`.xlsm` etc.) are ZIP archives. Inside, `xl/vbaProject.bin` (or `word/vbaProject.bin` for Word) is a "Compound Document" file, which is apparently the same binary format as the old `.xls` / `.doc` files.

vbaXray:

1. Copies the Office file to a temporary `.zip`, uses `Shell.Application` to extract `vbaProject.bin`, then opens it with `StgOpenStorage` (for file-based input) or `StgOpenStorageOnILockBytes` (for in-memory byte arrays).
2. Reads the `VBA/dir` stream and decompresses it using the MS-OVBA LZ77 variant (see MS-OVBA 2.4.1).
3. Parses the `dir` stream records to extract each module's name, stream name, and compressed-source offset.
4. Reads the `PROJECT` stream (plain text) to refine module types.
5. For each module, reads the corresponding VBA stream, slices from the stored offset, and decompresses the source.

All COM interfaces (`IStorage`, `IStream`, `IEnumSTATSTG`) use `DispCallFunc` vtable dispatch — this is primarily because I'm still learning how to make type libraries, but it has the benefit of being portable.

---

## API reference

### Methods

| Method | Description |
|---|---|
| `LoadFromFile(FilePath)` | Load from a macro-enabled Office file, or a `vbaProject.bin`. |
| `LoadFromByteArray(data())` | Load from a byte array already in memory (e.g. extracted from an office file). |
| `ExportAll(ExportPath)` | Export all modules to `ExportPath`, organised into numbered sub-folders by type. |
| `ExportModule(Index, ExportPath)` | Export a single module by 1-based index. |
| `ExportModuleByName(Name, ExportPath)` | Export a single module by name (case-insensitive). |
| `FullSourceCode([Separator])` | Returns the source of every module concatenated, with an optional header separator between each. The separator supports a `{0}` placeholder which is replaced with the module name. |
| `Reset()` | Clear all loaded data and return the object to its initial state. |

### Properties

| Property | Type | Description |
|---|---|---|
| `ProjectName` | `String` | The `Name=` value from the PROJECT stream. |
| `CodePage` | `Long` | The ANSI code page used to encode module source (e.g. 1252). |
| `IsPasswordProtected` | `Boolean` | `True` if the VBA project has a password set (beta) |
| `ModuleCount` | `Long` | Number of modules found in the project. |
| `ModuleName(Index)` | `String` | Name of the module at 1-based `Index`. |
| `ModuleType(Index)` | `ModuleType` | One of `modStandard`, `modClass`, `modDocument`, `modDesigner`, `modOther`. |
| `IsForm(Index)` | `Boolean` | `True` if the module is a UserForm designer module. |
| `SourceCode(Index)` | `String` | Full decompressed source of the module at `Index`. |

### `ModuleType` enum

```vba
modOther    = 0   ' Unknown / unclassified
modStandard = 1   ' .bas  — regular code module
modClass    = 2   ' .cls  — class module
modDesigner = 3   ' .frm  — UserForm or designer
modDocument = 4   ' .cls  — document module (eg. ThisWorkbook, Sheet1)
```

### Export folder structure

`ExportAll` writes files into the following sub-folders:

```
ExportPath\
    01_Standard\     ← .bas files
    02_Class\        ← .cls files
    03_Document\     ← .cls files (ThisWorkbook, Sheet1, …)
    04_Designer\     ← .frm files (UserForms)
    05_Other\        ← .txt fallback
```

---

## Limitations

- **Windows only** — relies on `kernel32` and `ole32`.
- **FRX files** — UserForm binary data (images, control properties etc.) stored in `.frx` files is not extracted. Only the `.frm` text stream is read, which contains the VBA source.
- **Read-only** — writing modified source back into the binary is not in scope for v1.
- **Does not export P-Code** — only the original VBA source. The P-Code streams are ignored for the simple reason that I have no idea how to decode them. I would recommend using looking at the OLEVBA/ViperMonkey projects and the work by BeakerBoy (https://github.com/Beakerboy/).
- **Access databases** (`.accdb` / `.accde`) are not supported.

---

## Resources + References

- [MS-OVBA specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/) — the authoritative reference for the VBA storage format.
- [Compound File Binary Format (MS-CFB)](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/)
- [GitHub - Excel-ZipTools](https://github.com/cristianbuse/excel-ziptools)
- [GitHub - Beakerboy/MS-OVBA](https://github.com/Beakerboy/MS-OVBA)

---

## License

MIT — see [LICENSE](LICENSE).

---

## Author

Kallun Willock — [https://github.com/KallunWillock](https://github.com/KallunWillock)
