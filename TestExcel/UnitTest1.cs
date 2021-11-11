using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestExcel
{
    [TestClass]
    public class UnitTest1
    {

        public static Cell temp = new Cell();
        public static Dictionary<string, Cell> dict = new Dictionary<string, Cell>() { { "A1", temp } };
        public Calculator calculator = new Calculator(dict);
        string dataLost = "";
        [TestMethod()]
        public void EvaluateTest1()          // тестування зчитування виразу
        {
            string expr = "(-2*5)^2-(-mod(27,10))";
            double expectedRes = 107;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvalueteTest2()        // тестування зчитування виразу
        {
            string expr = "-div(14, 5)*5^2 + 25/5 - mod(7, 4)^2";
            double expectedRes = -54;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void DivideZeroTest()      // тестування ділення на нуль
        {
            string expr = "5/0";  //вираз
            double res1 = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(dict["A1"].Val, "Error");
        }

        [TestMethod()]
        public void EvalueteTest3()        // тестування зчитування виразу
        {
            string expr = "2 + 5 - 5";
            double expectedRes = 2;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }
        [TestMethod()]
        public void EvalueteTest4()        // тестування зчитування виразу
        {
            string expr = "(-100)^2 -(-50)";
            double expectedRes = 10050;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }
        [TestMethod()]
        public void EvalueteTest5()        // тестування зчитування виразу
        {
            string expr = "div((5 + 2), 4)";
            double expectedRes = 1;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvalueteTest6()        // тестування зчитування виразу
        {
            string expr = "-(-1) + (-1)";
            double expectedRes = 0;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvalueteTest7()        // тестування зчитування виразу
        {
            string expr = "(50/5*5)^2";
            double expectedRes = 2500;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvalueteTest8()        // тестування зчитування виразу
        {
            string expr = "(-div(15, 4))^2";
            double expectedRes = 9;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvalueteTest9()        // тестування зчитування виразу
        {
            string expr = "2^2^2";
            double expectedRes = 16;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }


    }
}
