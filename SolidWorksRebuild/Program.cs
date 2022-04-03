using System;
using static System.Environment;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class App
{
    static SldWorks sldWorks = null!;

    public static void Main(string[] args)
    {
        sldWorks = new SldWorks();

        ModelDoc2? doc = sldWorks.ActiveDoc as ModelDoc2;
        if(doc is null)
            AppExit("Нет открытого документа.");

        doc!.Extension.Rebuild((int)swRebuildOptions_e.swForceRebuildAll);
    }

    #region Helpers
   
    static void WL(string text) => Console.WriteLine(text);

    static void AppExit(string text, int exitCode = -1)
    {
        WL(text);
        Exit(exitCode);
    }
    #endregion

}