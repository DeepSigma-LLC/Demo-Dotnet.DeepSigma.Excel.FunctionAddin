using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace Demo_DeepSigma.Excel.FunctionAddin;

/// <summary>
/// This class is the entry point for the Excel add-in. 
/// It implements the IExcelAddIn interface, which allows us to run code when the add-in is loaded and unloaded.
/// </summary>
public class AddIn : IExcelAddIn
{
    /// <summary>
    /// This method is called when the add-in is loaded.
    /// We use it to install the IntelliSense server, which provides enhanced IntelliSense support for our custom functions.
    /// </summary>
    public void AutoOpen()
    {
        IntelliSenseServer.Install();
    }
    
    /// <summary>
    /// This method is called when the add-in is unloaded.
    /// We use it to uninstall the IntelliSense server, which cleans up any resources used by the server.
    /// </summary>
    public void AutoClose()
    {
        IntelliSenseServer.Uninstall();
    }
}
