#region Imports

using Microsoft.VisualStudio.TestTools.UnitTesting;

using VBTranspiler.Parser;

#endregion

namespace VBTranspiler.CodeGenerator.UnitTests
{
  [TestClass]
  public class TestFormUserControlCodeGenerator : TestBase
  {
    protected override CodeGeneratorBase CreateCodeGenerator(VisualBasic6Parser.ModuleContext parseTree)
    {
      return new UserControlCodeGenerator(parseTree);
    }

    [TestMethod]
    public void TestClassNameTakenFromVBNameAttributeForUserControl()
    {
      string inputCode =
@"VERSION 1.0 CLASS
Attribute VB_Name = ""SomeUserControl""
";

      string expectedCode =
@"Imports System
Imports System.Windows.Forms
Imports Microsoft.VisualBasic

Public Class SomeUserControl
    Inherits UserControl

End Class
";

      VerifyGeneratedCode(inputCode, expectedCode);
    }
  }
}
