using ExcelDna.Integration;

namespace RExcelBridge;

public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        RBridge.TryStart();
    }

    public void AutoClose()
    {
        RBridge.Stop();
    }
}
