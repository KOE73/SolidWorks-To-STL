using Microsoft.Extensions.Configuration;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using static System.Environment;

public static class App
{
    static SldWorks sldWorks = null!;
    static string OutPath = String.Empty;

    public static void Main(string[] args)
    {
        sldWorks = new SldWorks();

        Microsoft.Extensions.Configuration.IConfiguration configuration = new ConfigurationBuilder()
            .AddCommandLine(args, new Dictionary<string, string>()
            {
                { "-H", "Help" },
                { "-O", "Out" }
            })
            .Build();

        var cmdHelp = configuration["Help"];
        bool setHelp = !string.IsNullOrWhiteSpace(cmdHelp);
        if(setHelp || args.Contains("-h") || args.Contains("-H"))
            ShowHelp();

        OutPath = configuration["Out"] ?? String.Empty;

        WL($"OutPath:{OutPath}");
        WL(String.Empty);

        Process();
    }

    private static void Process()
    {
        ModelDoc2? doc = sldWorks.ActiveDoc as ModelDoc2;
        if(doc is null)
            AppExit("Нет открытого документа.");

        ProcessDoc(doc!);
    }


    static void ProcessDoc(ModelDoc2 doc)
    {
        var docType = (swDocumentTypes_e)doc.GetType();

        switch(docType)
        {
            case swDocumentTypes_e.swDocASSEMBLY: ProcessAssemply(doc); break;
            case swDocumentTypes_e.swDocPART: ProcessPart(doc); break;
        }
    }

    static void ProcessAssemply(ModelDoc2 doc)
    {
        AssemblyDoc assemblyDoc = (AssemblyDoc)doc;

        var components = ((object[])assemblyDoc.GetComponents(false)).OfType<IComponent2>();
        var pathNames = components.Select(o => o.GetPathName()).Distinct();
        WL($"Components count:{pathNames.Count()}");

        foreach(var pathName in pathNames)
        {
            var innerDocSpec = (DocumentSpecification)sldWorks.GetOpenDocSpec(pathName);

            innerDocSpec.Silent = true;
            var innerDoc = sldWorks.OpenDoc7(innerDocSpec);

            int err = 0;

            var title = innerDoc.GetTitle();
            WL($"Title:{title}");
            ModelDoc2 swRefModel2 = (ModelDoc2)sldWorks.ActivateDoc3(title, false, (int)swRebuildOnActivation_e.swUserDecision, ref err);

            innerDoc.ViewZoomtofit2();

            BeginIndent();
            ProcessDoc(swRefModel2);
            //ProcessPart(swRefModel2);
            EndIndent();

            sldWorks.CloseDoc(pathName);
        }
    }

    static void ProcessPart(ModelDoc2 doc)
    {
        foreach(string configName in (string[])doc.GetConfigurationNames())
        {
            // Select configuration
            doc.ShowConfiguration2(configName);

            // Find SolidBodyFolder
            var feature = ((object[])doc.FeatureManager.GetFeatures(true)).OfType<Feature>().First(o => o.GetTypeName2() == "SolidBodyFolder");
            var bodyFolder = (BodyFolder)feature.GetSpecificFeature2();

            // Print info
            var config = doc.ConfigurationManager.ActiveConfiguration;
            WL($"Config:{config.Name} - {config.Description}\tBody сount:{bodyFolder.GetBodyCount()}");

            BeginIndent();
            foreach(Body2 body in (object[])bodyFolder.GetBodies())
            {
                WL($"Body name:{body.Name}\tFace count:{body.GetFaceCount()}");

                // Select all faces in the body
                var selectData = ((ISelectionMgr)doc.SelectionManager).CreateSelectData();
                var faces = (object[])body.GetFaces();
                doc.Extension.MultiSelect2(ObjectArrayToDispatchWrapper(faces), false, selectData);

                // Save. File name as body name!
                Save(doc, fileName: body.Name);
            }
            EndIndent();
        }
    }

    static void Save(IModelDoc2 doc, string fileName)
    {
        fileName = Path.ChangeExtension(fileName, ".STL");
        var fullName = Path.Combine(OutPath, fileName);

        WL($"Save:{fullName}");

        var longstatus = doc.SaveAs3(fullName,
            (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
            (int)swSaveAsOptions_e.swSaveAsOptions_Copy);
    }

    #region Helpers
    static int Indent = 0;
    static string IndentStr = String.Empty;

    static void BeginIndent()
    {
        Indent++;
        IndentStr = new('\t', Indent);
    }

    static void EndIndent()
    {
        Indent--;
        if(Indent < 0)
            Indent = 0;
        IndentStr = new('\t', Indent);
    }

    static void WL(string text) => Console.WriteLine(IndentStr + text);

    static DispatchWrapper[] ObjectArrayToDispatchWrapper(IEnumerable<object> objects) => objects.Select(o => new DispatchWrapper(o)).ToArray();

    static void ShowHelp()
    {
        WL($"Out      -O  Output path.");
        Exit(-1);
    }

    static void AppExit(string text, int exitCode = -1)
    {
        WL(text);
        Exit(exitCode);
    }
    #endregion
}
