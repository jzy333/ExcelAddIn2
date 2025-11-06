# ExcelAddIn2

Excel VSTO Add?in targeting .NET Framework4.8. It uses Ribbon XML for a simple custom tab and adds a left?docked Explorer pane. Workbooks are constrained to a single worksheet containing one structured table.

Features
- Custom ribbon (Ribbon XML): `My Add?in` tab with a `Say Hello` button.
- Explorer pane: a `CustomTaskPane` hosting a tree view of open Workbooks ? Worksheets ? Tables (ListObjects) ? Named Ranges.
- Structured layout enforcement:
 - One worksheet only.
 - Clears the sheet and creates a `ListObject` table named `DataTable` with columns `ID`, `Name`, `Position`.
 - Limits selection to the table area (`ScrollArea`) and protects workbook structure to prevent adding/removing sheets.

Requirements
- Windows with Microsoft Excel (Desktop).
- Visual Studio with Office developer tools (VSTO) and .NET Framework4.8 targeting pack.

Build and Run (Development)
- Open the solution in Visual Studio and press F5. Excel launches with the add?in loaded.
- Match Office bitness:
 -64?bit Office ? set project `Platform target` = x64.
 -32?bit Office ? set `Platform target` = x86.

Publish (ClickOnce)
- Right?click project ? `Publish` ? choose Folder or ClickOnce profile.
- Distribute the publish folder; users install by running the `.vsto` file.

Repository layout
- `ExcelAddIn2/ThisAddIn.cs`: Add?in startup, task pane, event wiring, and single?table enforcement.
- `ExcelAddIn2/RibbonXml.cs`: `[ComVisible(true)]` Ribbon XML handler and callbacks.
- `ExcelAddIn2/Ribbon.xml`: Optional embedded/file Ribbon XML.
- `ExcelAddIn2/WorkbookExplorerPane.cs`: Explorer `UserControl` with `TreeView`.

CI (GitHub Actions)
- Workflow builds the project on Windows and uploads `bin/Release` artifacts.

Troubleshooting
- Ribbon doesn’t appear: ensure `CreateRibbonExtensibilityObject()` returns a `RibbonXml` instance and that `RibbonXml` is `[ComVisible(true)]`.
- Add?in disabled: Excel ? File ? Options ? Trust Center ? Trust Center Settings ? Add?ins. Also check COM Add?ins dialog.
- Table not enforced: confirm Office bitness matches project platform target; events `WorkbookOpen`/`NewWorkbook` fire only for this instance of Excel.
