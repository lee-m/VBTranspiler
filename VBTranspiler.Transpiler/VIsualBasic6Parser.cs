#region Imports

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Antlr4.Runtime;

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
      return parser.module();
    }
  }
}
