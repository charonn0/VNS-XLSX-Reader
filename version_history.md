# Version History

This project follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

While `MAJOR=0`, breaking changes can occur on `MINOR` bumps.

## [0.1.0] - 2026-05-10

### Added

- Pure-Xojo XLSX reader in shared `Common/` folder, referenced from both `.xojo_project` files via `Module=`/`Class=` entries pointing at `../Common/`. Cross-platform: Mac / Windows / Linux desktop, plus Web 2.0 server runtime.
- Single-call public API: `XLSXReader.Open(file As FolderItem) As XLSXWorkbook` + a `MemoryBlock` overload for in-memory inputs.
- Workbook model: `XLSXWorkbook` (sheets / sharedStrings / styles), `XLSXSheet` (sparse cell dictionary + merged ranges), `XLSXCell` (typed value + lazy `DisplayText(styles)`), `XLSXCellRange`, `XLSXCellRef`.
- Format engine: `XLSXFormatter` covers `General` / `0` / `0.00` / `#,##0` / `#,##0.00` / `0%` / `0.00%` for numbers and `yyyy-mm-dd` / `dd/mm/yyyy` / `hh:mm` / `hh:mm:ss` / `yyyy-mm-dd hh:mm` for dates. Excel epoch 1899-12-30. Built-in numFmt ids 0..22 seeded in `XLSXStyles`; custom `<numFmts>` codes override.
- Zip part reader (`XLSXZip`) using the framework's own `FolderItem.Unzip`; magic-byte verification; per-instance temp-folder cleanup in `Destructor`. Both `FolderItem` and `MemoryBlock` entry points.
- Typed exception (`XLSXException`) with `eParseError` code (`NotAZip` / `MissingPart` / `MalformedXML` / `Encrypted` / `Unsupported`) plus an `XLSXHelpers` global module with `Extends`-based `eEnum.ToString`.
- Desktop app: `MainWindow` with `DesktopTabPanel TabPanelSheets` + `DesktopListBox ListboxData`; `MainMenuBar` File → Open (Cmd-O) opens a filtered `OpenFileDialog`; per-sheet tabs; column auto-sizing + horizontal scrollbar; user-drag column resizing.
- Web app: `MainPage` with `WebFileUploader UploaderXLSX` (auto-triggers `StartUpload` on `FileAdded`) + `WebTabPanel TabPanelSheets` + `WebListBox ListboxData`; per-sheet tabs; same UX as Desktop.
- Both fillers skip styled-but-empty rows so workbooks like `BUA 2024` (485 nominal rows, mostly empty) render only their real content.
- Localizable user-visible strings in the `strings` module; UI errors mapped to localizable kStrError… constants.
- Developer documentation: [`developper_doc.md`](developper_doc.md).
- Test fixtures: 4 public XLSX samples in [`test_files/`](test_files/).

### Verified manually

- Both projects analyze cleanly in Xojo 2026r1 with no API 2 warnings.
- Desktop opens a 33-sheet test workbook end-to-end: tabs populated, listbox repopulates on tab change, shared-string lookup resolves, merged-cell anchors render correctly, columns auto-sized + user-resizable + horizontally scrollable.
- Web opens the same workbook end-to-end via file upload.

### Known limitations

- Read-only; no formula evaluation (cached values shown); no images / charts / pivots / conditional formatting; no encrypted workbooks; cell colors / fonts / borders are not rendered.
- Parser is eager — large workbooks (≫ 10k rows × 50 cols) may load slowly.
- Format-code support is a pragmatic subset; unknown codes fall back to `Double.ToString` / `DateTime.SQLDateTime`.
