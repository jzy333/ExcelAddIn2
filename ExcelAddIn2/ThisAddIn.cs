using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn2
{
    public partial class ThisAddIn
    {
        private WorkbookExplorerPane _explorer;
        private Microsoft.Office.Tools.CustomTaskPane _pane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
#if DEBUG
            System.Diagnostics.Debug.WriteLine("ThisAddIn_Startup reached");
#endif
            // Ensure Explorer pane
            if (_explorer == null)
            {
                _explorer = new WorkbookExplorerPane();
                _pane = this.CustomTaskPanes.Add(_explorer, "Explorer");
                _pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
                _pane.Width =300;
                _pane.Visible = true;
            }

            HookEvents();

            // Apply enforcement to all currently open workbooks
            try
            {
                foreach (Excel.Workbook wb in this.Application.Workbooks)
                {
                    EnsureWorkbookLayout(wb);
                }
            }
            catch { }

            RefreshExplorer();
        }

        private void HookEvents()
        {
            try
            {
                this.Application.WorkbookOpen += Application_WorkbookOpen;
                ((Excel.AppEvents_Event)this.Application).NewWorkbook += new Excel.AppEvents_NewWorkbookEventHandler(Application_NewWorkbook);
                this.Application.WorkbookActivate += Application_WorkbookActivate;
                this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
                this.Application.SheetActivate += Application_SheetActivate;
            }
            catch { }
        }

        private void UnhookEvents()
        {
            try
            {
                this.Application.WorkbookOpen -= Application_WorkbookOpen;
                ((Excel.AppEvents_Event)this.Application).NewWorkbook -= new Excel.AppEvents_NewWorkbookEventHandler(Application_NewWorkbook);
                this.Application.WorkbookActivate -= Application_WorkbookActivate;
                this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
                this.Application.SheetActivate -= Application_SheetActivate;
            }
            catch { }
        }

        private void Application_NewWorkbook(Excel.Workbook Wb)
        {
            EnsureWorkbookLayout(Wb);
            RefreshExplorer();
        }

        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            EnsureWorkbookLayout(Wb);
            RefreshExplorer();
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            RefreshExplorer();
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            RefreshExplorer();
        }

        private void Application_SheetActivate(object Sh)
        {
            RefreshExplorer();
        }

        private void RefreshExplorer()
        {
            if (_explorer != null)
            {
                _explorer.RefreshTree(this.Application);
            }
        }

        // Ensure the workbook has a single sheet with a table (ID, Name, Position) and prevent adding sheets
        private void EnsureWorkbookLayout(Excel.Workbook wb)
        {
            if (wb == null) return;
            try
            {
                bool oldAlerts = this.Application.DisplayAlerts;
                this.Application.DisplayAlerts = false;
                try
                {
                    // Unprotect structure if needed (no password assumed)
                    try { wb.Unprotect(Type.Missing); } catch { }

                    // Keep first worksheet; remove others
                    Excel.Worksheet first = wb.Worksheets.Count >=1 ? (Excel.Worksheet)wb.Worksheets[1] : null;
                    if (first == null)
                    {
                        // Create one if none exists (shouldn't happen)
                        first = (Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing,1, Type.Missing);
                    }
                    for (int i = wb.Worksheets.Count; i >=1; i--)
                    {
                        Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[i];
                        if (!object.ReferenceEquals(ws, first))
                        {
                            try { ws.Delete(); } catch { }
                        }
                    }

                    // Clear sheet and remove any existing tables
                    first.Cells.Clear();
                    try
                    {
                        foreach (Excel.ListObject lo in first.ListObjects)
                        {
                            try { lo.Delete(); } catch { }
                        }
                    }
                    catch { }

                    // Headers
                    var headers = new[] { "ID", "Name", "Position" };
                    for (int c =0; c < headers.Length; c++)
                    {
                        first.Cells[1, c +1].Value2 = headers[c];
                    }

                    // Create table using header row plus one empty data row (Excel requires a rectangular range)
                    Excel.Range tableRange = first.Range[first.Cells[1,1], first.Cells[2, headers.Length]];
                    Excel.ListObject tbl = first.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tableRange, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing);
                    tbl.Name = "DataTable";
                    try { tbl.TableStyle = "TableStyleMedium2"; } catch { }
                    first.Columns.AutoFit();

                    // Restrict selection to the table only (no extra ranges)
                    try
                    {
                        string address = tbl.Range.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                        first.ScrollArea = address;
                    }
                    catch { }

                    // Protect workbook structure to prevent adding/removing sheets
                    wb.Protect(Type.Missing, true, Type.Missing);
                }
                finally
                {
                    this.Application.DisplayAlerts = oldAlerts;
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                System.Diagnostics.Debug.WriteLine("EnsureWorkbookLayout failed: " + ex);
#endif
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            UnhookEvents();
            if (_pane != null)
            {
                this.CustomTaskPanes.Remove(_pane);
                _pane = null;
            }
            _explorer = null;
        }

        // Return a custom IRibbonExtensibility implementation so the runtime loads our Ribbon XML
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
#if DEBUG
            System.Diagnostics.Debug.WriteLine("CreateRibbonExtensibilityObject called");
#endif
            return new RibbonXml();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
