using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using VBTranspiler.Parser;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace VBTranspiler.CodeGenerator
{
  public abstract class FormsCodeGeneratorBase: ClassModuleCodeGenerator
  {
    public FormsCodeGeneratorBase(VisualBasic6Parser.ModuleContext parseTree) : base(parseTree)
    {
    }

    protected abstract TypeSyntax InheritsType { get;  }

    protected override TypeBlockSyntax CreateTopLevelTypeDeclaration(IEnumerable<StatementSyntax> members)
    {
      ClassBlockSyntax classDecl = (ClassBlockSyntax)base.CreateTopLevelTypeDeclaration(members);

      TypeSyntax[] typeArr = new TypeSyntax[] { InheritsType };
      InheritsStatementSyntax[] inheritArr = new InheritsStatementSyntax[] { SyntaxFactory.InheritsStatement(typeArr) };
      SyntaxList<InheritsStatementSyntax> inherits = SyntaxFactory.List(inheritArr);

      classDecl = classDecl.WithInherits(inherits);

      return classDecl;
    }

    protected override void AddAdditionalImports(List<ImportsStatementSyntax> imports)
    {
      imports.Add(CreateImportStatement("System.Windows.Forms"));
    }
  }

  public class FormCodeGenerator : FormsCodeGeneratorBase
  {
    public FormCodeGenerator(VisualBasic6Parser.ModuleContext parseTree) : base(parseTree)
    { }

    protected override TypeSyntax InheritsType
    { get { return SyntaxFactory.ParseTypeName("Form"); } }
  }

  public class UserControlCodeGenerator : FormsCodeGeneratorBase
  {
    public UserControlCodeGenerator(VisualBasic6Parser.ModuleContext parseTree) : base(parseTree)
    { }

    protected override TypeSyntax InheritsType
    { get { return SyntaxFactory.ParseTypeName("UserControl"); } }
  }
}
