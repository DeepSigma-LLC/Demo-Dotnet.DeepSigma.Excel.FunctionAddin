using Xunit;
using Demo_DeepSigma.Excel.FunctionAddin;

namespace DeepSigma.Excel.FunctionAddin.Test;

public class MyExcelFunctionsTest
{
    [Fact]
    public void TestAddTwo()
    {
        double result = MyExcelFunctions.AddTwo(2, 3);
        Assert.Equal(5, result);
    }

    [Fact]
    public void TestSayHello()
    {
        string result = MyExcelFunctions.SayHello("Alice");
        Assert.Equal("Hello Alice", result);
    }

    [Fact]
    public void TestReturnArray()
    {
        object[,] result = MyExcelFunctions.ReturnArray();
        Assert.Equal(1, result[0, 0]);
        Assert.Equal(2, result[0, 1]);
        Assert.Equal(3, result[0, 2]);
    }
}