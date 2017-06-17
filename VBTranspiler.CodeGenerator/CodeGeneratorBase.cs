#region Imports

using System;
using System.IO;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

using VBTranspiler.Parser;

#endregion

namespace VBTranspiler.CodeGenerator
{
  public abstract class CodeGeneratorBase : VisualBasic6BaseVisitor<CodeGeneratorBase>
  {
    /// <summary>
    /// The parse tree to generate code from.
    /// </summary>
    private VisualBasic6Parser.ModuleContext mParseTree;

    /// <summary>
    /// Members of the main class/module.
    /// </summary>
    private List<StatementSyntax> mMainDeclMembers;

    /// <summary>
    /// Initialises this instance with the given VB6 AST.
    /// </summary>
    /// <param name="parseTree"></param>
    public CodeGeneratorBase(VisualBasic6Parser.ModuleContext parseTree)
    {
      mParseTree = parseTree;
      mMainDeclMembers = new List<StatementSyntax>();
    }

    /// <summary>
    /// Generate the .NET code for a VB6 AST.
    /// </summary>
    /// <returns></returns>
    public string GenerateCode()
    {
      //Walk the parse tree, generating the equivalent .NET code along the way.
      Visit(mParseTree);

      //Build the list of imports
      List<ImportsStatementSyntax> imports = new List<ImportsStatementSyntax>();
      imports.Add(CreateImportStatement("System"));
      
      AddAdditionalImports(imports);
      imports.Add(CreateImportStatement("Microsoft.VisualBasic"));

      //Build the main compilation unit now that all the members have been processed.
      CompilationUnitSyntax cu = SyntaxFactory.CompilationUnit()
                                 .AddImports(imports.ToArray())
                                 .AddMembers(CreateTopLevelTypeDeclaration(mMainDeclMembers))
                                 .NormalizeWhitespace();

      return cu.ToFullString();
    }

    /// <summary>
    /// Creates the top level class or module declaration.
    /// </summary>
    /// <returns>The type declaration.</returns>
    protected abstract TypeBlockSyntax CreateTopLevelTypeDeclaration(IEnumerable<StatementSyntax> members);

    protected abstract void AddAdditionalImports(List<ImportsStatementSyntax> imports);

    /// <summary>
    /// Creates a module declaration.
    /// </summary>
    /// <returns>The module declaration AST node.</returns>
    ModuleBlockSyntax CreateModuleDeclaration()
    {
      SyntaxTokenList publicModifier = RoslynUtils.PublicModifier;
      return SyntaxFactory.ModuleBlock(SyntaxFactory.ModuleStatement(GetVBNameAttributeValue()).WithModifiers(publicModifier))
             .WithMembers(SyntaxFactory.List(mMainDeclMembers));
    }

    /// <summary>
    /// Creates a new import clause for the specified namespace.
    /// </summary>
    /// <param name="namespaceName">Namespace to import</param>
    /// <returns>The new import statement AST node</returns>
    protected ImportsStatementSyntax CreateImportStatement(string namespaceName)
    {
      var importsClause = SyntaxFactory.SimpleImportsClause(SyntaxFactory.ParseName(namespaceName));
      return SyntaxFactory.ImportsStatement(SyntaxFactory.SeparatedList<ImportsClauseSyntax>(new ImportsClauseSyntax[] { importsClause }));
    }

    /// <summary>
    /// Gets the name of the VB6 class/module
    /// </summary>
    /// <returns></returns>
    protected string GetVBNameAttributeValue()
    {
      foreach (VisualBasic6Parser.AttributeStmtContext attr in mParseTree.moduleAttributes().attributeStmt())
      {
        if (attr.implicitCallStmt_InStmt().GetText() == "VB_Name")
          return attr.literal()[0].STRINGLITERAL().GetText().Trim(new char[] { '"' });
      }

      throw new ApplicationException("Unable to determine class name.");
    }

    /// <summary>
    /// Generates .NET code for an enumeration. 
    /// </summary>
    /// <param name="context">Enumeration AST.</param>
    /// <returns></returns>
    public override CodeGeneratorBase VisitEnumerationStmt(VisualBasic6Parser.EnumerationStmtContext context)
    {
      SyntaxTokenList accessibility = new SyntaxTokenList();
      VisualBasic6Parser.PublicPrivateVisibilityContext vis = context.publicPrivateVisibility();

      if (vis != null)
      {
        if (vis.PUBLIC() != null)
          accessibility = RoslynUtils.PublicModifier;
        else if (vis.PRIVATE() != null)
          accessibility = RoslynUtils.PrivateModifier;
      }

      EnumStatementSyntax stmt = SyntaxFactory.EnumStatement(context.ambiguousIdentifier().GetText()).WithModifiers(accessibility);
      EnumBlockSyntax enumBlock = SyntaxFactory.EnumBlock(stmt);
      List<EnumMemberDeclarationSyntax> members = new List<EnumMemberDeclarationSyntax>();

      foreach (var constant in context.enumerationStmt_Constant())
      {
        string constantName = constant.ambiguousIdentifier().GetText();

        if (constant.valueStmt() != null)
        {
          string constantValue = constant.valueStmt().GetText();

          members.Add(SyntaxFactory.EnumMemberDeclaration(constantName)
                      .WithInitializer(SyntaxFactory.EqualsValue(SyntaxFactory.NumericLiteralExpression(SyntaxFactory.ParseToken(constantValue)))));

        }
        else
          members.Add(SyntaxFactory.EnumMemberDeclaration(constantName));
      }

      mMainDeclMembers.Add(enumBlock.WithMembers(SyntaxFactory.List<StatementSyntax>(members)));
      return this;
    }

    /// <summary>
    /// Generates .NET code a constant field decl.
    /// </summary>
    /// <param name="context">Const decl AST.</param>
    /// <returns></returns>
    public override CodeGeneratorBase VisitConstStmt(VisualBasic6Parser.ConstStmtContext context)
    {
      List<SyntaxToken> modifiers = new List<SyntaxToken>();
      VisualBasic6Parser.PublicPrivateGlobalVisibilityContext vis = context.publicPrivateGlobalVisibility();

      if (vis != null)
      {
        if (vis.PUBLIC() != null || vis.GLOBAL() != null)
          modifiers.Add(SyntaxFactory.Token(SyntaxKind.PublicKeyword));
        else if (vis.PRIVATE() != null)
          modifiers.Add(SyntaxFactory.Token(SyntaxKind.PrivateKeyword));
      }

      modifiers.Add(SyntaxFactory.Token(SyntaxKind.ConstKeyword));

      foreach (VisualBasic6Parser.ConstSubStmtContext subStmt in context.constSubStmt())
      {
        SimpleAsClauseSyntax asClause = null;

        if (subStmt.asTypeClause() != null)
          asClause = SyntaxFactory.SimpleAsClause(SyntaxFactory.ParseTypeName(subStmt.asTypeClause().type().GetText()));

        ModifiedIdentifierSyntax identifier = SyntaxFactory.ModifiedIdentifier(subStmt.ambiguousIdentifier().GetText());
        ExpressionSyntax initialiser = null;

        string initialiserExpr = subStmt.valueStmt().GetText();

        //ParseExpression can't handle date/time literals - Roslyn bug? - //so handle them manually here
        if (initialiserExpr.StartsWith("#"))
        {
          string trimmedInitialiser = initialiserExpr.Trim(new char[] { '#' });
          DateTime parsedVal = new DateTime();

          if (trimmedInitialiser.Contains("AM") || trimmedInitialiser.Contains("PM"))
            parsedVal = DateTime.ParseExact(trimmedInitialiser, "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture);
          else
            parsedVal = DateTime.ParseExact(trimmedInitialiser, "M/d/yyyy", CultureInfo.InvariantCulture);

          initialiser = SyntaxFactory.DateLiteralExpression(SyntaxFactory.DateLiteralToken(initialiserExpr, parsedVal));
        }
        else
          initialiser = SyntaxFactory.ParseExpression(initialiserExpr);

        VariableDeclaratorSyntax varDecl = SyntaxFactory.VariableDeclarator(identifier).WithInitializer(SyntaxFactory.EqualsValue(initialiser));

        if (asClause != null)
          varDecl = varDecl.WithAsClause(asClause);

        mMainDeclMembers.Add(SyntaxFactory.FieldDeclaration(varDecl).WithModifiers(SyntaxFactory.TokenList(modifiers)));
      }

      return base.VisitConstStmt(context);
    }
  }
}
