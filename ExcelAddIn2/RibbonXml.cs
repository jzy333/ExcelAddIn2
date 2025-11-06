using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace ExcelAddIn2
{
 [ComVisible(true)]
 public sealed class RibbonXml : IRibbonExtensibility
 {
 private IRibbonUI ribbon;

 public string GetCustomUI(string ribbonID)
 {
#if DEBUG
 System.Diagnostics.Debug.WriteLine($"GetCustomUI called for: {ribbonID}");
#endif
 // Prefer embedded resource if available
 var asm = Assembly.GetExecutingAssembly();
 using (var stream = asm.GetManifestResourceStream("ExcelAddIn2.Ribbon.xml"))
 {
 if (stream != null)
 {
 using (var reader = new StreamReader(stream))
 {
 return reader.ReadToEnd();
 }
 }
 }

 // Try loading from add-in folder if present
 try
 {
 var folder = Path.GetDirectoryName(asm.Location);
 var filePath = folder == null ? null : Path.Combine(folder, "Ribbon.xml");
 if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
 {
 return File.ReadAllText(filePath);
 }
 }
 catch { /* ignore and fall back to inline */ }

 // Fallback: inline Ribbon XML (hello button only)
 return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnLoad"">
 <ribbon>
 <tabs>
 <tab id=""MyCustomTab"" label=""My Add-in"" insertAfterMso=""TabDeveloper"">
 <group id=""MyGroup"" label=""Actions"">
 <button id=""btnHello"" label=""Say Hello"" size=""large"" onAction=""OnHelloClicked""/>
 </group>
 </tab>
 </tabs>
 </ribbon>
</customUI>";
 }

 public void OnLoad(IRibbonUI ribbonUI)
 {
 this.ribbon = ribbonUI;
 }

 public void OnHelloClicked(IRibbonControl control)
 {
 System.Windows.Forms.MessageBox.Show("Hello from Ribbon XML");
 }
 }
}
