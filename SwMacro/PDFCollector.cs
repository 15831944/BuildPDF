using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;

namespace BuildPDF.csproj {
  class PDFCollector {
    private List<FileInfo> lfi = new List<FileInfo>();
    private List<KeyValuePair<string, string>> nf = new List<KeyValuePair<string, string>>();
    private DrawingData d = new DrawingData();

    public PDFCollector(SldWorks swApp) {
      _swApp = swApp;
    }

    public void Collect() {
      string fullpath = (SwApp.ActiveDoc as ModelDoc2).GetPathName();
      FileInfo top_level = d.GetPath(Path.GetFileNameWithoutExtension(fullpath));
      lfi.Add(top_level);

      collect_drwgs((ModelDoc2)SwApp.ActiveDoc);
    }

    private void collect_drwgs(ModelDoc2 md) {
      SWTableType swt = null;
      string title = md.GetTitle();
      try {
        swt = new SWTableType(md, Properties.Settings.Default.TableHashes);
      } catch (Exception e) {
        System.Diagnostics.Debug.WriteLine(e.Message);
      }

      List<FileInfo> ss = new List<FileInfo>();
      if (swt != null) {
        string part = string.Empty;
        bool in_lfi;
        bool in_nf;
        for (int i = 1; i < swt.RowCount; i++) {
          System.Diagnostics.Debug.WriteLine("table: " + swt.GetProperty(i, "PART NUMBER"));
          part = swt.GetProperty(i, "PART NUMBER");
          if (!part.StartsWith("0")) {
            FileInfo fi = d.GetPath(part);
            in_lfi = is_in(part, lfi);
            in_nf = is_in(part, nf);
            if (fi != null) {
              if (!in_lfi) {
                ss.Add(fi);
              } else {
                break;
              }
            } else {
              if (!in_nf) {
                nf.Add(new KeyValuePair<string, string>(part, title));
              } else {
                break;
              }
            }
          } else {
            System.Diagnostics.Debug.WriteLine("Skipping " + part);
          }
        }

        lfi.AddRange(ss);
      }

      if (ss.Count > 0) {
        foreach (FileInfo f in ss) {
          if (f != null) {
            string doc = f.FullName.ToUpper().Replace(@"K:\", @"G:\").Replace(".PDF", ".SLDDRW");
            SwApp.OpenDoc(doc, (int)swDocumentTypes_e.swDocDRAWING);
            SwApp.ActivateDoc(doc);
            ModelDoc2 m = (ModelDoc2)SwApp.ActiveDoc;
            System.Diagnostics.Debug.WriteLine("ss   : " + f.Name);
            System.Diagnostics.Debug.WriteLine(doc);
            collect_drwgs(m);
            SwApp.CloseDoc(doc);
          }
        }
      }
    }

    public static bool is_in(FileInfo f, List<FileInfo> l) {
      foreach (FileInfo fi in l) {
        if (f != null && Path.GetFileNameWithoutExtension(f.Name).ToUpper() == 
          Path.GetFileNameWithoutExtension(fi.Name).ToUpper()) {
          return true;
        }
      }
      return false;
    }

    public static bool is_in(FileInfo f, List<KeyValuePair<string, string>> l) {
      foreach (KeyValuePair<string, string> fi in l) {
        if (f != null && Path.GetFileNameWithoutExtension(f.Name).ToUpper() == 
    Path.GetFileNameWithoutExtension(fi.Key).ToUpper()) {
          return true;
        }
      }
      return false;
    }

    public static bool is_in(string f, List<FileInfo> l) {
      foreach (FileInfo fi in l) {
        if (f != null && f.ToUpper() == Path.GetFileNameWithoutExtension(fi.Name).ToUpper()) {
          return true;
        }
      }
      return false;
    }

    public static bool is_in(string f, List<KeyValuePair<string, string>> l) {
      foreach (KeyValuePair<string, string> fi in l) {
        if (f != null && f.ToUpper() == Path.GetFileNameWithoutExtension(fi.Key).ToUpper()) {
          return true;
        }
      }
      return false;
    }

    public List<FileInfo> PDFCollection {
      get { return lfi; }
      set { lfi = value; }
    }

    public List<KeyValuePair<string, string>> NotFound {
      get { return nf; }
      set { nf = value; }
    }

    private SldWorks _swApp;

    public SldWorks SwApp {
      get { return _swApp; }
      private set { _swApp = value; }
    }

  }
}
