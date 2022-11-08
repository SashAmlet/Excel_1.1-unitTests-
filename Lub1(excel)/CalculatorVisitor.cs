using Antlr4.Runtime.Misc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lub1_excel_
{
    internal class CalculatorVisitor : CalculatorBaseVisitor<double>
    {
        public override double VisitCompileUnit(CalculatorParser.CompileUnitContext context)
        {
            return Visit(context.expr());
        }
        public override double VisitNum(CalculatorParser.NumContext context)
        {
            var result = double.Parse(context.GetText());
            return result;
        }
        private void minusProblem(Cell cell)
        {
            string[] slova = cell.strValue.Split('-');
            string lit = null;

            if (slova.Length > 1)
            {
                cell.Formula = slova[0];
                for (int ii = 1; ii < slova.Length; ii++)
                {
                    foreach (char c in slova[ii])
                    {
                        if (c == '(' || c == '+' || c == '*' || c == '/' || c == ':' || c == '%')
                            break;
                        lit += c;
                    }
                    if (lit == null)
                        cell.Formula += string.Format("-{0}", slova[ii]);
                    else
                    {
                        cell.Formula += string.Format("-({0})", lit);
                        char[] chLit = lit.ToCharArray();
                        cell.Formula += slova[ii].TrimStart(chLit);
                        lit = null;
                    }
                }
                cell.Formula = cell.Formula.TrimStart('=');
            }
            else
                cell.Formula = cell.strValue.TrimStart('=');
        }
        public override double VisitIdent(CalculatorParser.IdentContext context)
        {
            var result = context.GetText();
            if (CellManager.Instance.outOfArrCheck())
                throw new IndexOutOfRangeException();
            Cell cell = CellManager.Instance.GetCell(result);
            
            CellManager.Instance[CellManager.Instance.index++] = cell;

            if (cell.strValue == null)
                return 0;

            if (cell.doubValue != 0)
                return cell.doubValue;

            // for making (-5)+3 from -5+3
            minusProblem(cell);

            double doubValue = Calculator.Evaluate(cell.Formula);
            return doubValue;
        }
        public override double VisitParent(CalculatorParser.ParentContext context)
        {
            return Visit(context.expr());
        }
        public override double VisitNeg(CalculatorParser.NegContext context)
        {
            var right = Visit(context.right);
            return -right;
        }
        public override double VisitAddSub(CalculatorParser.AddSubContext context)
        {
            var left = Visit(context.left);
            var right = Visit(context.right);
            if (context.op.Type == CalculatorLexer.ADD)
            {
                return left + right;
            }
            else
            {
                return (left - right);
            }
        }
        public override double VisitMulDiv(CalculatorParser.MulDivContext context)
        {
            var left = Visit(context.left);
            var right = Visit(context.right);
            if (context.op.Type == CalculatorLexer.MULTIPLY)
            {
                return left * right;
            }
            else if (right == 0)
                throw new DivideByZeroException();
            else if (context.op.Type == CalculatorLexer.DIVIDE)
            {
                return left / right;
            }
            else if (context.op.Type == CalculatorLexer.DIV)
            {
                return ((int)(left / right));
            }
            else
            {
                return left % right;
            }
        }
        public override double VisitCompare(CalculatorParser.CompareContext context)
        {
            var left = Visit(context.left);
            var right = Visit(context.right);
            if (context.op.Type == CalculatorLexer.LESS)
            {
                if (double.Parse(left.ToString()) < double.Parse(right.ToString()))
                    return 1;
                return 0;

            }
            else if (context.op.Type == CalculatorLexer.LESSOREQ)
            {
                if (double.Parse(left.ToString()) <= double.Parse(right.ToString()))
                    return 1;
                return 0;

            }
            else if (context.op.Type == CalculatorLexer.GREATER)
            {
                if (double.Parse(left.ToString()) > double.Parse(right.ToString()))
                    return 1;
                return 0;
            }
            else if (context.op.Type == CalculatorLexer.GREATOREQ)
            {
                if (double.Parse(left.ToString()) >= double.Parse(right.ToString()))
                    return 1;
                return 0;

            }
            else
            {
                if (double.Parse(left.ToString()) == double.Parse(right.ToString()))
                    return 1;
                return 0;
            }

        }
    }
}
