#region Imports

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#endregion

namespace VBTranspiler.CodeGenerator.UnitTests
{
  [TestClass]
  public class TestCodeGenerationEnum : TestBase
  {
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
      VerifyGeneratedCode(inputCode, expectedCode, RoslynCodeGenerator.SourceType.Class);
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
      VerifyGeneratedCode(inputCode, expectedCode, RoslynCodeGenerator.SourceType.Class);
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
      VerifyGeneratedCode(inputCode, expectedCode, RoslynCodeGenerator.SourceType.Class);
    }
  }
}
