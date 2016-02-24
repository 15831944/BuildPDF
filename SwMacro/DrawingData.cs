using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data.SqlClient;

namespace BuildPDF.csproj {
  class DrawingData {
    SqlConnection sqc = new SqlConnection();
    public DrawingData() {
      sqc = new SqlConnection(Properties.Settings.Default.CutlistConnectionString);

      try {
        sqc.Open();
      } catch (InvalidOperationException ioe) {
        throw new Exception(@"The connection is already open, or information is missing from the connection string: '" +
            Properties.Settings.Default.CutlistConnectionString + "'.", ioe);
      } catch (SqlException se) {
        throw new Exception(@"A connection-level error occurred while opening the connection, whatever that means.", se);
      } catch (Exception e) {
        throw new Exception(@"Whoops. There's a problem", e);
      }

    }

    public FileInfo GetPath(string filename) {
      string SQL = "SELECT FPath + FName AS FullPath FROM GEN_DRAWINGS WHERE FName LIKE @fname";
      using (SqlCommand comm = new SqlCommand(SQL, sqc)) {
        comm.Parameters.AddWithValue("@fname", filename + '%');
        using (SqlDataReader sdr = comm.ExecuteReader()) {
          if (sdr.Read()) {
            return new FileInfo(sdr.GetString(0));
          }
        }
      }
      return null;
    }

  }
}
