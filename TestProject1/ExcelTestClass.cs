using Lub1_excel_;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Assert = NUnit.Framework.Assert;

namespace TestProject1
{
    [TestFixture]
    public class ExcelTestClass
    {
        [Test]
        public void addSubTest()
        {
            Assert.That(Calculator.Evaluate("6-(3+9)-(4)+5"), Is.EqualTo(-5));
        }
        [Test]
        public void mulDivTest()
        {
            Assert.That(Calculator.Evaluate("15/3*4*2/5"), Is.EqualTo(8));
        }
        [Test]
        public void compareTest()
        {
            Assert.That(Calculator.Evaluate("15>3"), Is.EqualTo(1));
        }

        [Test]
        public void multyTest()
        {
            char[] sign = { '+', '-', '*', '/' };
            Random random = new Random();
            string expression = string.Empty + '1';
            long expValue = 1;
            int a, b;

            for (int i = 0; i < 10; i++)
            {
                a = random.Next(1, 100);
                b = random.Next(0, sign.Length - 1);
                expression = expression.Insert(0, "(");
                expression = expression + sign[b] + /*(sign[b] == '-' ?*/ '(' + a.ToString() + ')' + ')'/*: a.ToString() + ')') */;
                switch (sign[b])
                {
                    case '+':
                        expValue += a;
                        break;
                    case '-':
                        expValue -= a;
                        break;
                    case '*':
                        expValue *= a;
                        break;
                    case '/':
                        expValue /= a;
                        break;
                    default:
                        break;
                }

            }
            Assert.That(Calculator.Evaluate(expression), Is.EqualTo(expValue));
        }
    }
}
