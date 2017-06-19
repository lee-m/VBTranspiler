#region Imports

using System;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Antlr4.Runtime;
using Antlr4.Runtime.Tree;

#endregion

namespace VBTranspiler.Parser
{
  public partial class VisualBasic6Parser
  {
    public static ModuleContext ParseSource(string fileName)
    {
      using(FileStream stream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
      {
        return ParseSource(stream);
      }
    }

    public static ModuleContext ParseSource(Stream stm)
    {
      AntlrInputStream input = new AntlrInputStream(stm);
      ITokenStream tokens = new CommonTokenStream(new VisualBasic6Lexer(input));

      VisualBasic6Parser parser = new VisualBasic6Parser(tokens);
      parser.AddParseListener(new ParserListener(parser));
      parser.AddErrorListener(new DebugErrorListener<IToken>());

      var ret = parser.module();

      if (parser.NumberOfSyntaxErrors > 0)
        throw new ApplicationException("Parser errors encountered");

      return ret;
    }
  }

  public class ParserListener : IParseTreeListener
  {
    private int mIndent;
    private VisualBasic6Parser mParser;

    public ParserListener(VisualBasic6Parser parser)
    {
      mParser = parser;
    }

    public void EnterEveryRule(ParserRuleContext ctx)
    {
      mIndent += 1;

      Debug.Write("".PadLeft(mIndent));
      Debug.WriteLine(string.Format("Enter {0} {1}", mParser.RuleNames[ctx.RuleIndex], ctx.Start.Text));
    }

    public void ExitEveryRule(ParserRuleContext ctx)
    {
      Debug.Write("".PadLeft(mIndent));
      Debug.WriteLine(string.Format("Exit {0}", mParser.RuleNames[ctx.RuleIndex])); 
      
      mIndent -= 1;
    }

    public void VisitErrorNode(IErrorNode node)
    {
    }

    public void VisitTerminal(ITerminalNode node)
    {
    }
  }

  public class DebugErrorListener<Symbol> : IAntlrErrorListener<Symbol>
  {
    public virtual void SyntaxError(IRecognizer recognizer, Symbol offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
    {
      Debug.WriteLine("line " + line + ":" + charPositionInLine + " " + msg);
    }
  }
}
