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

      int l = pc.PDFCollection.Count;
      bool last = false;
      m.Append("Merging ");
      foreach (FileInfo s in pc.PDFCollection) {
        if (s != null) {
          last = l-- < 2;
          m.Append((last ? "and " : string.Empty) + s.Name + (last ? ".\n" : ", "));
          m.Refresh();
        }
      }

      m.Append("\n");

      foreach (KeyValuePair<string, string> n in pc.NotFound) {
        m.AppendLine(string.Format("No drawing for '{0}' in '{1}'.",
          n.Key,
          n.Value.Split(new string[] { " - " }, StringSplitOptions.None)[0].Trim()));
      }

      System.GC.Collect(0, GCCollectionMode.Forced);

      string tmpPath = Path.GetTempFileName().Replace(".tmp", ".PDF");
      string path = Properties.Settings.Default.TargetPath + pc.PDFCollection[0].Name;

      PDFMerger pm = new PDFMerger(pc.PDFCollection, new FileInfo(tmpPath));
      pm.Merge();

      try {
        File.Copy(tmpPath, path, true);
      } catch (Exception e) {
        m.AppendLine(e.Message);
      }

      m.AppendLine("Created '" + path + "'.");
      m.AppendLine("Opening...");
      System.Diagnostics.Process.Start(path);
      System.GC.Collect(0, GCCollectionMode.Forced);
    }
	
    /// <summary>
    ///  The SldWorks swApp variable is pre-assigned for you.
    /// </summary>
    public SldWorks swApp;
  }
}


