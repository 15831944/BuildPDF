using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

using iTextSharp.text;
using iTextSharp.text.pdf;

namespace BuildPDF.csproj {
  class PDFMerger {
    public PDFMerger(List<FileInfo> lfi, FileInfo target) {
      _target = target;
      _pdf_paths.AddRange(lfi);
    }

    public void Merge() {
      byte[] ba = merge_files(_pdf_paths);
      using (FileStream fs = File.Create(_target.FullName)) {
        for (int i = 0; i < ba.Length; i++) {
          fs.WriteByte(ba[i]);
        }
      }
    }

    public static byte[] merge_files(List<FileInfo> docs) {
      Document document = new Document();
      using (MemoryStream ms = new MemoryStream()) {
        PdfCopy copy = new PdfCopy(document, ms);
        document.Open();
        int document_page_counter = 0;

        foreach (FileInfo fi in docs) {
          PdfReader reader = new PdfReader(fi.FullName);

          for (int i = 1; i <= reader.NumberOfPages; i++) {
            document_page_counter++;
            PdfImportedPage ip = copy.GetImportedPage(reader, i);
#if PAGE_NUMBERS
            PdfCopy.PageStamp ps = copy.CreatePageStamp(ip);

            ColumnText.ShowTextAligned(ps.GetOverContent(), Element.ALIGN_CENTER,
              new Phrase(string.Format("{0}", document_page_counter)), (ip.Width / 2), 5, ip.Width < ip.Height ? 0 : 1);
            ps.AlterContents();
#endif
            copy.AddPage(ip);
          }

          copy.FreeReader(reader);
          reader.Close();
        }
        document.Close();
        return ms.GetBuffer();
      }
    }

    //private List<Document> _pdfs = new List<Document>();

    //public List<Document> PDFs {
    //  get { return _pdfs; }
    //  set { _pdfs = value; }
    //}

    private FileInfo _target;

    public FileInfo Target {
      get { return _target; }
      set { _target = value; }
    }


    private List<FileInfo> _pdf_paths = new List<FileInfo>();

    public List<FileInfo> PDFCollection {
      get { return _pdf_paths; }
      set { _pdf_paths = value; }
    }

  }
}
