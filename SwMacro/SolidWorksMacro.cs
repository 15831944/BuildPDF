using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;
using System.IO;
using System.Collections.Generic;

namespace BuildPDF.csproj {
  public partial class SolidWorksMacro {

    public void Main() {
      Message m = new Message();
      PDFCollector pc = new PDFCollector(swApp);
      m.Show();
      m.AppendLine("Collecting PDF paths...");
      pc.Collect();

      foreach (FileInfo s in pc.PDFCollection) {
        if (s != null) {
          m.Append("Merging " + s.Name + ", ");
          m.Refresh();
        }
      }

      m.Append("\n");

      foreach (KeyValuePair<string, string> n in pc.NotFound) {
        m.AppendLine(string.Format("No drawing for '{0}' in '{1}'.",
          n.Key,
          n.Value.Split(new string[] { " - " }, StringSplitOptions.None)[0].Trim()));
      }

      string tmpPath = Properties.Settings.Default.TargetPath + pc.PDFCollection[0].Name;

      PDFMerger pm = new PDFMerger(pc.PDFCollection, new FileInfo(tmpPath));
      pm.Merge();

      m.AppendLine("Created " + tmpPath);
      m.AppendLine("Opening...");
      System.Diagnostics.Process.Start(tmpPath);
      System.GC.Collect(0, GCCollectionMode.Forced);
    }
	
    /// <summary>
    ///  The SldWorks swApp variable is pre-assigned for you.
    /// </summary>
    public SldWorks swApp;
  }
}


