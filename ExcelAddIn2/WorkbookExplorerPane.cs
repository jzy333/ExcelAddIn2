using System;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn2
{
 public class WorkbookExplorerPane : UserControl
 {
 private readonly TreeView _tree;

 public WorkbookExplorerPane()
 {
 _tree = new TreeView
 {
 Dock = DockStyle.Fill,
 HideSelection = false
 };
 Controls.Add(_tree);
 }

 public void RefreshTree(Excel.Application app)
 {
 if (app == null) return;

 _tree.BeginUpdate();
 try
 {
 _tree.Nodes.Clear();

 foreach (Excel.Workbook wb in app.Workbooks)
 {
 var wbNode = new TreeNode(wb.Name) { Tag = wb }; 

 // Worksheets
 var sheetsNode = new TreeNode("Worksheets");
 foreach (Excel.Worksheet ws in wb.Worksheets)
 {
 var wsNode = new TreeNode(ws.Name) { Tag = ws };

 // Tables in worksheet (ListObjects)
 var tablesNode = new TreeNode("Tables");
 try
 {
 foreach (Excel.ListObject lo in ws.ListObjects)
 {
 tablesNode.Nodes.Add(new TreeNode(lo.Name) { Tag = lo });
 }
 }
 catch { /* ignore COM errors */ }

 // Named ranges on worksheet
 var namesNode = new TreeNode("Named Ranges");
 try
 {
 foreach (Excel.Name name in ws.Names)
 {
 namesNode.Nodes.Add(new TreeNode(name.Name) { Tag = name });
 }
 }
 catch { /* ignore COM errors */ }

 if (tablesNode.Nodes.Count >0) wsNode.Nodes.Add(tablesNode);
 if (namesNode.Nodes.Count >0) wsNode.Nodes.Add(namesNode);

 sheetsNode.Nodes.Add(wsNode);
 }
 if (sheetsNode.Nodes.Count >0) wbNode.Nodes.Add(sheetsNode);

 // Workbook-level named ranges
 var wbNamesNode = new TreeNode("Workbook Named Ranges");
 try
 {
 foreach (Excel.Name name in wb.Names)
 {
 // Skip sheet-scoped names duplicated at workbook level if desired
 wbNamesNode.Nodes.Add(new TreeNode(name.Name) { Tag = name });
 }
 }
 catch { /* ignore */ }
 if (wbNamesNode.Nodes.Count >0) wbNode.Nodes.Add(wbNamesNode);

 _tree.Nodes.Add(wbNode);
 }

 if (_tree.Nodes.Count >0)
 {
 _tree.Nodes[0].Expand();
 }
 }
 finally
 {
 _tree.EndUpdate();
 }
 }
 }
}
