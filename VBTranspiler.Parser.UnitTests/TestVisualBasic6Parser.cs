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

      Assert.AreEqual("BorderStyle", controlProperties[0].cp_SingleProperty().ambiguousIdentifier().GetText());
      Assert.AreEqual("3", controlProperties[0].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("Caption", controlProperties[1].cp_SingleProperty().ambiguousIdentifier().GetText());
      Assert.AreEqual(@"""Some Form""", controlProperties[1].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("ClientHeight", controlProperties[2].cp_SingleProperty().ambiguousIdentifier().GetText());
      Assert.AreEqual("7950", controlProperties[2].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("MaxButton", controlProperties[3].cp_SingleProperty().ambiguousIdentifier().GetText());
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
         ItemData        =   ""frmJobDetails.frx"":0000
         Height          =   315
         Left            =   9600
         TabIndex        =   3
         Top             =   780
         Width           =   330
         _Version        =   21563
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

      Assert.AreEqual("BorderStyle", controlProperties[0].cp_SingleProperty().ambiguousIdentifier().GetText());
      Assert.AreEqual("3", controlProperties[0].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("Caption", controlProperties[1].cp_SingleProperty().ambiguousIdentifier().GetText());
      Assert.AreEqual(@"""Some Form""", controlProperties[1].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("ClientHeight", controlProperties[2].cp_SingleProperty().ambiguousIdentifier().GetText());
      Assert.AreEqual("7950", controlProperties[2].cp_SingleProperty().literal().GetText());

      Assert.AreEqual("MaxButton", controlProperties[3].cp_SingleProperty().ambiguousIdentifier().GetText());
      Assert.AreEqual("0", controlProperties[3].cp_SingleProperty().literal().GetText());

      //Frame nested control block
      var frameControlBlock = controlProperties[4].controlProperties();

      Assert.AreEqual(7, frameControlBlock.cp_Properties().Length);
      Assert.AreEqual("VB.Frame", frameControlBlock.cp_ControlType().GetText());
      Assert.AreEqual("SomeFrame", frameControlBlock.cp_ControlIdentifier().GetText());

      //Button nested control block
      var buttonControlBlock = frameControlBlock.cp_Properties().Last().controlProperties();

      Assert.AreEqual(8, buttonControlBlock.cp_Properties().Length);
      Assert.AreEqual("VB.CommandButton", buttonControlBlock.cp_ControlType().GetText());
      Assert.AreEqual("SomeButton", buttonControlBlock.cp_ControlIdentifier().GetText());

      var frxOffset = buttonControlBlock.cp_Properties()[1].cp_SingleProperty().cp_FrxOffset();
      Assert.AreEqual("0000", frxOffset.literal().GetText()); 
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
      EndProperty
      Top             =   120
      Width           =   10755
  End
End
";
      var parseTree = ParseInputSource(inputSource);
      var moduleControls = parseTree.controlProperties();
    }
  }
}
