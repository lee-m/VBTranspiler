#region Imports

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using VBTranspiler.Parser;

#endregion

namespace VBTranspiler.CodeGenerator.UnitTests
{
  [TestClass]
  public class TestModuleCodeGenerator : TestBase
  {
    protected override CodeGeneratorBase CreateCodeGenerator(VisualBasic6Parser.ModuleContext parseTree)
    {
      return new ModuleCodeGenerator(parseTree);
    }

    [TestMethod]
    public void TestClassNameTakenFromVBNameAttributeForModule()
    {
      string inputCode =
@"VERSION 1.0 CLASS
Attribute VB_Name = ""SomeModule""
";

      string expectedCode =
@"Imports System
Imports Microsoft.VisualBasic

Public Module SomeModule
End Module
";

      VerifyGeneratedCode(inputCode, expectedCode);
    }
  }
}
