using Antlr4.Runtime;
using Lub1_excel_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lub1_excel_
{
    public static class Calculator
    {
        public static double Evaluate(string expression)
        {
            var lexer = new CalculatorLexer(new AntlrInputStream(expression));
            lexer.RemoveErrorListeners();
            var tokens = new CommonTokenStream(lexer);
            var parser = new CalculatorParser(tokens);
            var tree = parser.compileUnit();
            var visitor = new CalculatorVisitor();
            return visitor.Visit(tree);
        }
    }
}
