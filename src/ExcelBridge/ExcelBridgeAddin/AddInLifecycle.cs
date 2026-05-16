using ExcelDna.Integration;
using System;

public class AddInLifecycle : IExcelAddIn
{
    public void AutoOpen()
    {
        // Nothing required at startup. The worker is started lazily by CPP_PIPE_START
        // or by the first function that needs the pipe.
    }

    public void AutoClose()
    {
        // Excel-DNA calls this when the XLL is unloaded or Excel closes normally.
        // Never allow cleanup errors to interrupt Excel shutdown.
        try
        {
            PipeFunctions.CPP_PIPE_STOP();
        }
        catch
        {
            // Suppress all shutdown cleanup errors.
        }
    }
}
