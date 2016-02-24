using System;
using System.Collections.Generic;
using System.Text;

namespace BuildPDF.csproj {
  class Exceptions {

    [global::System.Serializable]
    public class BuildPDFException : Exception {
      //
      // For guidelines regarding the creation of new exception types, see
      //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconerrorraisinghandlingguidelines.asp
      // and
      //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dncscol/html/csharp07192001.asp
      //

      public BuildPDFException() { }
      public BuildPDFException(string message) : base(message) { }
      public BuildPDFException(string message, Exception inner) : base(message, inner) { }
      protected BuildPDFException(
      System.Runtime.Serialization.SerializationInfo info,
      System.Runtime.Serialization.StreamingContext context)
        : base(info, context) { }
    }
  }
}
