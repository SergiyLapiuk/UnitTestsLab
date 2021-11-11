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
        public void EvaluateTest1()          // ���������� ���������� ������
        {
            string expr = "(-2*5)^2-(-mod(27,10))";
            double expectedRes = 107;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvalueteTest2()        // ���������� ���������� ������
        {
            string expr = "-div(14, 5)*5^2 + 25/5 - mod(7, 4)^2";
            double expectedRes = -54;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void DivideZeroTest()      // ���������� ������ �� ����
        {
            string expr = "5/0";  //�����
            double res1 = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(dict["A1"].Val, "Error");
        }

        [TestMethod()]
        public void EvalueteTest3()        // ���������� ���������� ������
        {
            string expr = "2 + 5 - 5";
            double expectedRes = 2;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }
        [TestMethod()]
        public void EvalueteTest4()        // ���������� ���������� ������
        {
            string expr = "(-100)^2 -(-50)";
            double expectedRes = 10050;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }
        [TestMethod()]
        public void EvalueteTest5()        // ���������� ���������� ������
        {
            string expr = "div((5 + 2), 4)";
            double expectedRes = 1;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvalueteTest6()        // ���������� ���������� ������
        {
            string expr = "-(-1) + (-1)";
            double expectedRes = 0;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvalueteTest7()        // ���������� ���������� ������
        {
            string expr = "(50/5*5)^2";
            double expectedRes = 2500;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvalueteTest8()        // ���������� ���������� ������
        {
            string expr = "(-div(15, 4))^2";
            double expectedRes = 9;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvalueteTest9()        // ���������� ���������� ������
        {
            string expr = "2^2^2";
            double expectedRes = 16;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }


    }
}
