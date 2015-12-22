#region Imports

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#endregion

namespace VBTranspiler.CodeGenerator.UnitTests
{
  [TestClass]
  public class TestCodeGenerationFieldDecls : TestBase
  {
    [TestMethod]
    public void TestConstantFieldDeclCodeGeneration()
    {
      string inputCode =
@"VERSION 1.0 CLASS
Attribute VB_Name = ""SomeClass""

Private Const Constant1 As Integer = 77
Private Const Constant2 = """"
Const Constant3 As String = ""X"", Constant4 = 42
Const Constant5 = New Collection
Const Constant6 = #10/4/2015 3:23:00 AM#
Const Constant7 = #12/25/2014#
Const Constant8 = #1/1/2015 2:56:00 PM#
Const Constant9 = (5 + 1) / 2 * 3
";

      string expectedCode =
@"Imports System
Imports Microsoft.VisualBasic

Public Class SomeClass

    Private Const Constant1 As Integer = 77

    Private Const Constant2 = """"

    Const Constant3 As String = ""X""

    Const Constant4 = 42

    Const Constant5 = New Collection

    Const Constant6 = #10/4/2015 3:23:00 AM#

    Const Constant7 = #12/25/2014#

    Const Constant8 = #1/1/2015 2:56:00 PM#

    Const Constant9 =(5 + 1) / 2 * 3
End Class
";
      VerifyGeneratedCode(inputCode, expectedCode, RoslynCodeGenerator.SourceType.Class);
    }
  }
}
