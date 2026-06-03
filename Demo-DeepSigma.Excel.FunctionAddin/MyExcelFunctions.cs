using ExcelDna.Integration;

namespace Demo_DeepSigma.Excel.FunctionAddin;

/// <summary>
/// This class contains the custom functions that will be exposed to Excel.
/// </summary>
public static class MyExcelFunctions
{
    /// <summary>
    /// This function takes two numbers and returns their sum.
    /// </summary>
    /// <param name="x">The first number.</param>
    /// <param name="y">The second number.</param>
    /// <returns>The sum of the two numbers.</returns>
    [ExcelFunction(Description = "Adds two numbers", 
        Category = "DeepSigma", 
        HelpTopic="This is the help topic")]
    public static double AddTwo(double x, double y) => x + y;

    /// <summary>
    ///  This function takes a name as input and returns a greeting message.
    /// </summary>
    /// <param name="name">The name of the person to greet.</param>
    /// <returns>A greeting message.</returns>
    [ExcelFunction(Description = "Returns a greeting", 
        Category = "DeepSigma", 
        HelpTopic="This is the help topic")]
    public static string SayHello(
        [ExcelArgument(Name = "name", Description = "The name of the person to greet.")]
        string name) => "Hello " + name;


    /// <summary>
    /// This function returns a 2D array of numbers.
    /// </summary>
    /// <returns>A 2D array of numbers.</returns>
    [ExcelFunction(
        Description = "Returns an array of numbers", 
        Category = "DeepSigma", 
        HelpTopic="This is the help topic")]
    public static object[,] ReturnArray()
    {
        return new object[,]
        {
            { 1, 2, 3 }
        };
    }


    /// <summary>
    /// This function demonstrates how to use an ExcelArgument attribute to provide metadata for a function argument.
    /// </summary>
    /// <returns>A 2D array with mixed types, including strings, numbers, and Excel errors.</returns>
    [ExcelFunction(Description = "Returns a mixed output array", Category = "DeepSigma", HelpTopic="This is the help topic")]
    public static object[,] MixedOutput()
    {
        return new object[,]
        {
        { "Ticker", "Price", "Status" },
        { "MSFT", 425.10, "OK" },
        { "BAD", ExcelError.ExcelErrorNA, ExcelEmpty.Value }
        };
    }
}