#region Imports

using Microsoft.VisualStudio.TestTools.UnitTesting;

using VBTranspiler.Parser;

#endregion

namespace VBTranspiler.CodeGenerator.UnitTests
{
  [TestClass]
  public class TestFormCodeGenerator : TestBase
  {
    protected override CodeGeneratorBase CreateCodeGenerator(VisualBasic6Parser.ModuleContext parseTree)
    {
      return new FormCodeGenerator(parseTree);
    }

    [TestMethod]
    public void TestClassNameTakenFromVBNameAttributeForForm()
    {
      string inputCode =
@"VERSION 1.0 CLASS
Attribute VB_Name = ""SomeForm""
";

      string expectedCode =
@"Imports System
Imports System.Windows.Forms
Imports Microsoft.VisualBasic

Public Class SomeForm
    Inherits Form

End Class
";

      VerifyGeneratedCode(inputCode, expectedCode);
    }

   
  }
}
