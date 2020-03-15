using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime.Misc;

namespace Antonio_Spreadsheets
{
    class Visitor : LanguageBaseVisitor<double>
    {
        Dictionary<string, double> tableIdentifier = new Dictionary<string, double>();
        public override double VisitCompileUnit(LanguageParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitNumberExpr(LanguageParser.NumberExprContext context)
        {
            var result = double.Parse(context.GetText());
            Debug.WriteLine(result);

            return result;
        }

        //IdentifierExpr
        public override double VisitIdentifierExpr(LanguageParser.IdentifierExprContext context)
        {
            var result = context.GetText();
            double value;
            //видобути значення змінної з таблиці
            if (tableIdentifier.TryGetValue(result.ToString(), out value))
            {
                return value;
            }
            else
            {
                return 0.0;
            }
        }

        public override double VisitParenthesizedExpr(LanguageParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitIncDecExpr([NotNull] LanguageParser.IncDecExprContext context)
        {
            var number = WalkLeft(context);
            if (context.operatorToken.Type == LanguageLexer.INC)
            {
                Debug.WriteLine("inc( {0} )", number);
                return number + 1;
            }
            else
            {
                Debug.WriteLine("dec( {0} )", number);
                return number - 1;
            }
        }

        public override double VisitMaxMinExpr([NotNull] LanguageParser.MaxMinExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LanguageLexer.MAX)
            {
                if (left >= right) return left;
                else return right;
            }
            else
            {
                if (left <= right) return left;
                else return right;
            }
        }

        public override double VisitModDivExpr([NotNull] LanguageParser.ModDivExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LanguageLexer.MOD)
                return left % right;
            else
                return (int)left / (int)right;
        }

        public override double VisitExponentialExpr(LanguageParser.ExponentialExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            Debug.WriteLine("{0} ^ {1}", left, right);
            return System.Math.Pow(left, right);
        }

        public override double VisitMinusExpr(LanguageParser.MinusExprContext context)
        {
            var num = WalkLeft(context);
            return 0-num;
        }

        public override double VisitPlusExpr(LanguageParser.PlusExprContext context)
        {
            return WalkLeft(context);
        }

        public override double VisitAdditiveExpr(LanguageParser.AdditiveExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            if (context.operatorToken.Type == LanguageLexer.ADD)
            {
                Debug.WriteLine("{0} + {1}", left, right);
                return left + right;
            }
            else //subtract
            {
                Debug.WriteLine("{0} - {1}", left, right);
                return left - right;
            }
        }

        public override double VisitMultiplicativeExpr(LanguageParser.MultiplicativeExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            if (context.operatorToken.Type == LanguageLexer.MULTIPLY)
            {
                Debug.WriteLine("{0} * {1}", left, right);
                return left * right;
            }
            else //divide
            {
                Debug.WriteLine("{0} / {1}", left, right);
                return left / right;
            }
        }

        private double WalkLeft(LanguageParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<LanguageParser.ExpressionContext>(0));
        }

        private double WalkRight(LanguageParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<LanguageParser.ExpressionContext>(1));
        }
    }
}