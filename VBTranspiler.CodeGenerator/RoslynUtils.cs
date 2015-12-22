#region Imports

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

#endregion

namespace VBTranspiler.CodeGenerator
{
  public static class RoslynUtils
  {
    /// <summary>
    /// Creates a syntax token list representing a public modifier.
    /// </summary>
    public static SyntaxTokenList PublicModifier { get { return SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.PublicKeyword)); } }

    /// <summary>
    /// Creates a syntax token list representing a private modifier.
    /// </summary>
    public static SyntaxTokenList PrivateModifier { get { return SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.PrivateKeyword)); } }

    /// <summary>
    /// Creates a syntax token list representing a friend modifier.
    /// </summary>
    public static SyntaxTokenList FriendModifier { get { return SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.FriendKeyword)); } }
  }
}
