#region Imports

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using VBTranspiler.Parser;
using VBTranspiler.CodeGenerator;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Microsoft.CodeAnalysis.VisualBasic;

#endregion

namespace VBTranspiler.CodeGenerator.UnitTests
{
  public class TestBase
  {
    protected VisualBasic6Parser.ModuleContext ParseInputSource(string source)
    {
      using (MemoryStream memStm = new MemoryStream(Encoding.ASCII.GetBytes(source)))
      {
        return VisualBasic6Parser.ParseSource(memStm);
      }
    }

    protected void VerifyGeneratedCode(string inputCode, string expectedCode, RoslynCodeGenerator.SourceType type)
    {
      var codeGen = new RoslynCodeGenerator(ParseInputSource(inputCode), type);
      var generatedCode = codeGen.GenerateCode();

      Assert.AreEqual(expectedCode, generatedCode);

      //Make sure the generated code is syntactically valid
      var parsedCode = VisualBasicSyntaxTree.ParseText(generatedCode);
      Assert.IsFalse(parsedCode.GetCompilationUnitRoot().ContainsDiagnostics);
    }
  }
}
