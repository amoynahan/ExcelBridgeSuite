using ExcelDna.Integration;

namespace JuliaExcelBridge;

public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        JBridge.TryStart();
    }

    public void AutoClose()
    {
        JBridge.Stop();
    }
}
