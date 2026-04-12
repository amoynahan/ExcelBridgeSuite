using ExcelDna.Integration;

namespace PythonExcelBridge;

public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        PythonBridge.TryStart();
    }

    public void AutoClose()
    {
        PythonBridge.Stop();
    }
}
