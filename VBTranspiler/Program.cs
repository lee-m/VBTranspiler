#region Import

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using VBTranspiler.Parser;
using VBTranspiler.CodeGenerator;

#endregion

namespace VBTranspiler
{
  class Program
  {
    static void Main(string[] args)
    {
      var inputFile = @"Test.cls";
      var outputFile = @"Test.vb";

      var parseTree = VisualBasic6Parser.ParseSource(inputFile);
      var codeGen = new ClassModuleCodeGenerator(parseTree);

      using(StreamWriter output = new StreamWriter(outputFile))
      {
        output.Write(codeGen.GenerateCode());
      }
    }
  }
}
