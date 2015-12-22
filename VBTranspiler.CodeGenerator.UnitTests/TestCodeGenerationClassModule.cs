#region Imports

using System;
using System.IO;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using VBTranspiler.Parser;
using VBTranspiler.CodeGenerator;

using Microsoft.CodeAnalysis.VisualBasic;

#endregion

namespace VBTranspiler.CodeGenerator.UnitTests
{
  [TestClass]
  public class TestCodeGenerationClassModule : TestBase
  {
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

      VerifyGeneratedCode(inputCode, expectedCode, RoslynCodeGenerator.SourceType.Class);
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

      VerifyGeneratedCode(inputCode, expectedCode, RoslynCodeGenerator.SourceType.Module);
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

      VerifyGeneratedCode(inputCode, expectedCode, RoslynCodeGenerator.SourceType.Form);
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

      VerifyGeneratedCode(inputCode, expectedCode, RoslynCodeGenerator.SourceType.UserControl);
    }
  }
}
