#region Imports

using System;
using System.IO;
using System.Linq;
using System.Text;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using VBTranspiler.Parser;

#endregion

namespace VBTranspiler.Parser.UnitTests
{
  [TestClass]
  public class TestVisualBasic6Parser
  {
    protected VisualBasic6Parser.ModuleContext ParseInputSource(string source)
    {
      using (MemoryStream memStm = new MemoryStream(Encoding.ASCII.GetBytes(source)))
      {
        return VisualBasic6Parser.ParseSource(memStm);
      }
    }

    [TestMethod]
    public void TestParsingModuleReferences()
    {
      string inputSource = @"
VERSION 5.00
Object = ""{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0""; ""Comdlg32.ocx""
Object = ""{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0""; ""RICHTX32.OCX""";

      var parseTree = ParseInputSource(inputSource);
      Assert.IsNotNull(parseTree.moduleReferences());

      var references = parseTree.moduleReferences().moduleReference();
      Assert.IsNotNull(references);
      Assert.AreEqual(2, references.Length);

      Assert.AreEqual(@"""{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0""", references[0].moduleReferenceGUID().GetText());
      Assert.AreEqual(@"""Comdlg32.ocx""", references[0].moduleReferenceComponent().GetText());

      Assert.AreEqual(@"""{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0""", references[1].moduleReferenceGUID().GetText());
      Assert.AreEqual(@"""RICHTX32.OCX""", references[1].moduleReferenceComponent().GetText());
    }

    [TestMethod]
    public void TestParsingModuleConfig()
    {
      string inputSource = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""SomeClass""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False";

      var parseTree = ParseInputSource(inputSource);
      Assert.IsNotNull(parseTree.moduleConfig());

      var configElems =parseTree.moduleConfig().moduleConfigElement();

      Assert.AreEqual(1, configElems.Length);
      Assert.AreEqual("MultiUse", configElems[0].ambiguousIdentifier().GetText());
      Assert.AreEqual("-1", configElems[0].literal().GetText());
    }

    [TestMethod()]
    public void TestParsingSingleLevelControlBlock()
    {
      string inputSource = @"
VERSION 5.00
Object = ""{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0""; ""Comdlg32.ocx""
Object = ""{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0""; ""RICHTX32.OCX""
Begin VB.Form SomeForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ""Some Form""
   ClientHeight    =   7950
   MaxButton       =   0   'False
End
";

      var parseTree = ParseInputSource(inputSource);
      var moduleControls = parseTree.controlProperties();

      Assert.IsNotNull(moduleControls);
      Assert.AreEqual("VB.Form", moduleControls.cp_ControlType().GetText());
      Assert.AreEqual("SomeForm", moduleControls.cp_ControlIdentifier().GetText());

      var controlProperties = parseTree.controlProperties().cp_Properties();
      Assert.IsNotNull(controlProperties);
      Assert.AreEqual(4, controlProperties.Length);

      Assert.AreEqual("BorderStyle", controlProperties[0].cp_SingleProperty().cp_PropertyName().GetText());
      Assert.AreEqual("3", controlProperties[0].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("Caption", controlProperties[1].cp_SingleProperty().cp_PropertyName().GetText());
      Assert.AreEqual(@"""Some Form""", controlProperties[1].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("ClientHeight", controlProperties[2].cp_SingleProperty().cp_PropertyName().GetText());
      Assert.AreEqual("7950", controlProperties[2].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("MaxButton", controlProperties[3].cp_SingleProperty().cp_PropertyName().GetText());
      Assert.AreEqual("0", controlProperties[3].cp_SingleProperty().literal().GetText());
    }

    [TestMethod()]
    public void TestParsingNestedLevelControlBlock()
    {
      string inputSource = @"
VERSION 5.00
Object = ""{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0""; ""Comdlg32.ocx""
Object = ""{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0""; ""RICHTX32.OCX""
Begin VB.Form SomeForm
  BorderStyle     =   3  'Fixed Dialog
  Caption         =   ""Some Form""
  ClientHeight    =   7950
  MaxButton       =   0   'False
  Begin VB.Frame SomeFrame
      Caption         =   ""Frame""
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   10755
      Begin VB.CommandButton SomeButton
         Caption         =   ""Button""
         ItemData        =   ""SomeForm.frx"":0000
         Height          =   315
         Left            =   9600
         TabIndex        =   3
         Top             =   780
         Width           =   330
         _Version        =   21563
         TextRTF         =   $""SomeForm.frx"":008A
         RightMargin     =   1.31072e5
         CurCell.BeginLfDblClick=   0   'False
      End
  End
End
";

      var parseTree = ParseInputSource(inputSource);
      var moduleControls = parseTree.controlProperties();

      Assert.AreEqual("VB.Form", moduleControls.cp_ControlType().GetText());
      Assert.AreEqual("SomeForm", moduleControls.cp_ControlIdentifier().GetText());

      var controlProperties = parseTree.controlProperties().cp_Properties();
      Assert.AreEqual(5, controlProperties.Length);

      Assert.AreEqual("BorderStyle", controlProperties[0].cp_SingleProperty().cp_PropertyName().GetText());
      Assert.AreEqual("3", controlProperties[0].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("Caption", controlProperties[1].cp_SingleProperty().cp_PropertyName().GetText());
      Assert.AreEqual(@"""Some Form""", controlProperties[1].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("ClientHeight", controlProperties[2].cp_SingleProperty().cp_PropertyName().GetText());
      Assert.AreEqual("7950", controlProperties[2].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("MaxButton", controlProperties[3].cp_SingleProperty().cp_PropertyName().GetText());
      Assert.AreEqual("0", controlProperties[3].cp_SingleProperty().literal().GetText());

      //Frame nested control block
      var frameControlBlock = controlProperties[4].controlProperties();

      Assert.AreEqual(7, frameControlBlock.cp_Properties().Length);
      Assert.AreEqual("VB.Frame", frameControlBlock.cp_ControlType().GetText());
      Assert.AreEqual("SomeFrame", frameControlBlock.cp_ControlIdentifier().GetText());

      //Button nested control block
      var buttonControlBlock = frameControlBlock.cp_Properties().Last().controlProperties();

      Assert.AreEqual(11, buttonControlBlock.cp_Properties().Length);
      Assert.AreEqual("VB.CommandButton", buttonControlBlock.cp_ControlType().GetText());
      Assert.AreEqual("SomeButton", buttonControlBlock.cp_ControlIdentifier().GetText());

      var frxOffset = buttonControlBlock.cp_Properties()[1].cp_SingleProperty().FRX_OFFSET();
      Assert.AreEqual(":0000", frxOffset.GetText());

      var secondFrxOffset = buttonControlBlock.cp_Properties()[8].cp_SingleProperty().FRX_OFFSET();
      Assert.AreEqual(":008A", secondFrxOffset.GetText());
    }

    [TestMethod()]
    public void TestParsingBeginPropertyBlock()
    {
      string inputSource = @"
VERSION 5.00
Object = ""{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0""; ""Comdlg32.ocx""
Object = ""{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0""; ""RICHTX32.OCX""
Begin VB.Form SomeForm
  BorderStyle     =   3  'Fixed Dialog
  Begin VB.Frame SomeFrame
      Caption         =   ""Frame""
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628}
          Text        =   ""Description""
          Height      =   315
          Left        =   9600
          TabInd      =   3
          BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
              Name            =   ""Courier New""
              Size            =   9
              Charset         =   0
              Weight          =   400
              Underline       =   0   'False
              Italic          =   0   'False
              Strikethrough   =   0   'False
          EndProperty
          BeginProperty SomeNestedProp 
              Name            =   ""Tahoma""
              Size            =   8.25
              Charset         =   0
              Weight          =   400
              Underline       =   0   'False
              Italic          =   0   'False
              Object.Strikethrough   =   0   'False
          EndProperty
          BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
          EndProperty
      EndProperty
      Top             =   120
      Width           =   10755
  End
End
";
      var parseTree = ParseInputSource(inputSource);
      var moduleControls = parseTree.controlProperties();
      Assert.AreEqual(2, moduleControls.cp_Properties().Length);

      var frameProp = moduleControls.cp_Properties()[1];
      Assert.AreEqual("SomeFrame", frameProp.controlProperties().cp_ControlIdentifier().GetText());

      var firstNestedProp = frameProp.controlProperties().cp_Properties()[1];

      Assert.AreEqual("ColumnHeader", firstNestedProp.cp_NestedProperty().ambiguousIdentifier().GetText());
      Assert.IsNotNull(firstNestedProp.cp_NestedProperty());
      Assert.AreEqual(7, firstNestedProp.cp_NestedProperty().cp_Properties().Length);

      var secondNestedProp = firstNestedProp.cp_NestedProperty().cp_Properties()[4];

      Assert.AreEqual("Font", secondNestedProp.cp_NestedProperty().ambiguousIdentifier().GetText());
      Assert.IsNotNull(secondNestedProp.cp_NestedProperty());
      Assert.AreEqual(7, secondNestedProp.cp_NestedProperty().cp_Properties().Length);
    }

    [TestMethod()]
    public void TestParsingPropertyReturningObject()
    {
      string inputSource = @"
Private Property Get iLibraryNode_PopupMenu() As Object
Dim obValues as clsValues
End Property";

      //Previous versions of the grammer couldn't handle Object as a type so simply check for no parse errors.
      ParseInputSource(inputSource);
    }

    [TestMethod()]
    public void TestParsingWhileLoopWithBlankLinesAtBodyEnd()
    {
      string inputSource = @"
Private Sub Test()

  Dim obValues As Object
  
  While True
    
      obValues = Nothing
	  
  Wend
  
End Sub
";
      //Used to fail with a parse error.
      ParseInputSource(inputSource);
    }

    [TestMethod()]
    public void TestParsingOptionExplicitAfterMemberVariables()
    {
      string inputSource = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = ""Class1""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private A As String
Private B As String
Private C As String
";
      //Used to fail with a parse error.
      ParseInputSource(inputSource);
    }

    [TestMethod()]
    public void TestParsingOptionStatementAfterMemberVariableDeclartion()
    {
      string inputSource = @"
VERSION 1.0 CLASS
Private A As String
Private B As String

Option Explicit

Private C as String

Option Compare Text
Option Base 0
";
      //Used to fail with a parse error.
      ParseInputSource(inputSource);
    }

    [TestMethod()]
    public void TestParsingMemberVariableDeclartionAfterOptionStatement()
    {
      string inputSource = @"
VERSION 1.0 CLASS
Option Explicit

Private A As String
Private B As String
";
      //Used to fail with a parse error.
      ParseInputSource(inputSource);
    }

    [TestMethod()]
    public void TestParsingTypeHintsOnArgumentIdentifiers()
    {
      string inputSource = @"
VERSION 1.0 CLASS
Option Explicit

Public Function Foo(a$, b&, c!, d#, e@, f$()) As Boolean
End Function

Public Sub Foo(a$, b&, c!, d#, e@)
End Sub

Public Property Get Bar(a$)
End Property

Public Property Let Bar(a$)
End Property

Public Property Set Bar(a$)
End Property

Event SomeEvent(a$)

Declare Function Func Lib ""Foo.dll"" (a$)
";
      //Used to fail with a parse error.
      ParseInputSource(inputSource);
    }

    [TestMethod()]
    public void TestParsingTypeHintsOnNumericLiterals()
    {
       string inputSource = @"
VERSION 1.0 CLASS
Option Explicit

Public Sub Foo()

  Dim x As Variant
  x = 123&
  x = 123!
  x = 123#
  x = 123@

End Sub
";
      //Used to fail with a parse error.
      ParseInputSource(inputSource);
    }

    [TestMethod()]
    public void TestParsingTypeHintsOnForLoopIndexVar()
    {
      string inputSource = @"
VERSION 1.0 CLASS
Option Explicit

Public Sub Foo()

  Dim i%
  
  For i% = 1 To 10 Step -1
    frame.Top = i%
    DoEvents
  Next i%

End Sub
";
      //Used to fail with a parse error.
      ParseInputSource(inputSource);
    }
  }
}
