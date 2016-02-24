using System;
using System.Collections.Generic;
using System.Text;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace BuildPDF.csproj {
  class SWTableType {
    private ModelDoc2 part;
    private SelectionMgr swSelMgr;
    private ITableAnnotation swTable;
    private int _col_count = 0;
    private int _row_count = 0;
    private List<string> _cols = new List<string>();
    private List<string> _prts = new List<string>();
    private string _masterHash = "00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00";
    private System.Collections.Specialized.StringCollection _masterHashes;
    private string _part_column = "PART NUMBER";
    private bool initialated = false;

    public SWTableType(ModelDoc2 md, string tablehash) {
      _masterHash = tablehash;
      _masterHashes[0] = _masterHash;
      part = md;
      swSelMgr = (SelectionMgr)part.SelectionManager;
      if (part != null && swSelMgr != null) {
        BomFeature swBom = null;
        try {
          swBom = (BomFeature)swSelMgr.GetSelectedObject6(1, -1);
        } catch {
          // 
        }
        if (swBom != null) {
          fill_table(swBom);
        } else {
          find_bom();
        }
      }
    }

    public SWTableType(ModelDoc2 md, System.Collections.Specialized.StringCollection sc) {
      _masterHashes = sc;
      part = md;
      swSelMgr = (SelectionMgr)part.SelectionManager;
      if (part != null && swSelMgr != null) {
        BomFeature swBom = null;
        try {
          swBom = (BomFeature)swSelMgr.GetSelectedObject6(1, -1);
        } catch {
          // 
        }
        if (swBom != null) {
          fill_table(swBom);
        } else {
          find_bom();
        }
      }
    }

    private void fill_table(BomFeature bom) {
      _cols.Clear();
      _prts.Clear();
      swTable = (ITableAnnotation)bom.IGetTableAnnotations(1);
      part.ClearSelection2(true);

      _col_count = swTable.ColumnCount;
      _row_count = swTable.RowCount;
      for (int i = 0; i < _col_count; i++) {
        _cols.Add(swTable.get_DisplayedText(0, i));
      }

      int prtcol = get_column_by_name(_part_column);
      for (int i = 0; i < _row_count; i++) {
        _prts.Add(swTable.get_DisplayedText(i, prtcol));
      }
      initialated = true;
    }

    private int get_column_by_name(string prop) {
      if (!initialated) {
        for (int i = 0; i < _col_count; i++) {
          if (swTable.get_DisplayedText(0, i).Trim().ToUpper().Equals(prop.ToUpper())) {
            return i;
          }
        }
      } else {
        int count = 0;
        foreach (string s in _cols) {
          if (s.Trim().ToUpper().Equals(prop.Trim().ToUpper())) {
            return count;
          }
          count++;
        }
      }
      return -1;
    }

    private int get_row_by_partname(string prt) {
      if (!initialated) {
        int prtcol = get_column_by_name(_part_column);
        for (int i = 0; i < _row_count; i++) {
          if (swTable.get_DisplayedText(i, prtcol).Trim().ToUpper().Equals(prt.Trim().ToUpper())) {
            return i;
          }
        }
      }
      return _prts.IndexOf(prt);
    }

    private string get_property_by_part(string prt, string prop, string part_column_name) {
      int prtrow = get_row_by_partname(prt);
      int prpcol = get_column_by_name(prop);
      return (prpcol < 1 || prtrow < 1) ? string.Empty : swTable.get_DisplayedText(prtrow, prpcol);
    }

    private string get_property_by_part(int row, string prop, string part_column_name) {
      int prpcol = get_column_by_name(prop);
      return (prpcol < 1 || row < 1) ? string.Empty : swTable.get_DisplayedText(row, prpcol);
    }

    private void find_bom() {
      bool found = false;
      Feature feature = (Feature)part.FirstFeature();
      if (part != null) {
        while (feature != null) {
          if (feature.GetTypeName2().ToUpper() == "BOMFEAT") {
            feature.Select2(false, -1);
            BomFeature bom = (BomFeature)swSelMgr.GetSelectedObject6(1, -1);
            fill_table(bom);
            if (identify_table(_cols, _masterHashes)) {
              found = true;
              System.Diagnostics.Debug.WriteLine("Found a table.");
              break;
            }
          }
          feature = (Feature)feature.GetNextFeature();
        }
      }
      if (!found) {
        throw new Exceptions.BuildPDFException("I couldn't find the correct table.");
      }
    }

    private bool identify_table(List<string> table, System.Collections.Specialized.StringCollection tablehash) {
      bool match = false;
      string str = string.Empty;
      string[] ss = new string[table.Count];
      table.CopyTo(ss);
      System.Array.Sort(ss);
      foreach (string s in ss) {
        str += string.Format("{0}|", s.ToUpper());
      }

      System.IO.Stream columns = new System.IO.MemoryStream();
      columns.Write(System.Text.Encoding.UTF8.GetBytes(str), 0, str.Length - 1);

      string hash = System.BitConverter.ToString(System.Security.Cryptography.MD5.Create().ComputeHash(to_byte_array(str)));

      foreach (string h in tablehash) {
        match |= hash == h;
      }

      return match;
    }

    private byte[] to_byte_array(string s) {
      byte[] ba = new byte[s.Length];
      int count = 0;
      foreach (char c in s) {
        ba[count] = (byte)c;
        count++;
      }
      return ba;
    }

    public string GetProperty(string part, string prop) {
      return get_property_by_part(part, prop, _part_column);
    }

    public string GetProperty(int row, string prop) {
      return get_property_by_part(row, prop, _part_column);
    }

    public string PartColumn {
      get { return _part_column; }
      set { _part_column = value; }
    }


    public int ColumnCount {
      get { return _col_count; }
      set { _col_count = value; }
    }

    public int RowCount {
      get { return _row_count; }
      set { _row_count = value; }
    }

    public List<string> Columns {
      get { return _cols; }
      set { _cols = value; }
    }

    public List<string> Parts {
      get { return _prts; }
      set { _prts = value; }
    }

    public string MasterHash {
      get { return _masterHash; }
      set { _masterHash = value; }
    }

  }
}
