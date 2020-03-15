using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
namespace Antonio_Spreadsheets
{

    public static class Calculator
    {
        public static double Evaluate(string expression)
        {
            var lexer = new LanguageLexer(new AntlrInputStream(expression));
            lexer.RemoveErrorListeners();
            lexer.AddErrorListener(new ThrowExceptionErrorListener());

            var tokens = new CommonTokenStream(lexer);
            var parser = new LanguageParser(tokens);
            parser.AddErrorListener(new ThrowExceptionErrorListener());

            var tree = parser.compileUnit();
            var visitor = new Visitor();

            return visitor.Visit(tree);
        }
    }
}
