#region Imports

using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBTranspiler.Parser;

#endregion

namespace VBTranspiler.CodeGenerator.UnitTests
{
  [TestClass]
  public class TestClassModuleCodeGenerator : TestBase
  {
    protected override CodeGeneratorBase CreateCodeGenerator(VisualBasic6Parser.ModuleContext parseTree)
    {
      return new ClassModuleCodeGenerator(parseTree);
    }

    [TestMethod]
    public void TestClassNameTakenFromVBNameAttributeForClassModule()
    {
      string inputCode =
@"VERSION 1.0 CLASS
Attribute VB_Name = ""SomeClass""
";

      string expectedCode =
@"Imports System
Imports Microsoft.VisualBasic

Public Class SomeClass
End Class
";

      VerifyGeneratedCode(inputCode, expectedCode);
    }

    [TestMethod]
    public void TestPublicEnumCodeGeneration()
    {
      string inputCode =
@"VERSION 1.0 CLASS
Attribute VB_Name = ""SomeClass""

Public Enum SomeEnum
  SECT_STUDY_DETAILS = 0
  SECT_STUDY_DEFINITION 
End Enum";

      string expectedCode =
@"Imports System
Imports Microsoft.VisualBasic

Public Class SomeClass

    Public Enum SomeEnum
        SECT_STUDY_DETAILS = 0
        SECT_STUDY_DEFINITION
    End Enum
End Class
";
      VerifyGeneratedCode(inputCode, expectedCode);
    }

    [TestMethod]
    public void TestPrivateEnumCodeGeneration()
    {
      string inputCode =
@"VERSION 1.0 CLASS
Attribute VB_Name = ""SomeClass""

Private Enum SomeEnum
  SECT_STUDY_DETAILS = 0
  SECT_STUDY_DEFINITION 
End Enum";

      string expectedCode =
@"Imports System
Imports Microsoft.VisualBasic

Public Class SomeClass

    Private Enum SomeEnum
        SECT_STUDY_DETAILS = 0
        SECT_STUDY_DEFINITION
    End Enum
End Class
";
      VerifyGeneratedCode(inputCode, expectedCode);
    }

    [TestMethod]
    public void TestNoVisibilityEnumCodeGeneration()
    {
      string inputCode =
@"VERSION 1.0 CLASS
Attribute VB_Name = ""SomeClass""

Enum SomeEnum
  SECT_STUDY_DETAILS = 0
  SECT_STUDY_DEFINITION 
End Enum";

      string expectedCode =
@"Imports System
Imports Microsoft.VisualBasic

Public Class SomeClass

    Enum SomeEnum
        SECT_STUDY_DETAILS = 0
        SECT_STUDY_DEFINITION
    End Enum
End Class
";
      VerifyGeneratedCode(inputCode, expectedCode);
    }

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
      VerifyGeneratedCode(inputCode, expectedCode);
    }
  }
}
