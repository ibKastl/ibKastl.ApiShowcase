using System.Diagnostics;
using System.IO;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;

namespace ApiShowcase.EPLAN.Scripting.Demo
{
  // ReSharper disable once UnusedMember.Global
  public class DemoScriptingAction
  {
    [DeclareAction("DemoScriptingAction")]

    // ReSharper disable once UnusedMember.Global
    public void Action()
    {
      // Export data via label action
      string outputPath = @"C:\Test\";

      string schemePartList = "Summarized parts list";
      string filePartList = Path.Combine(outputPath, "Summarized parts list.txt");
      Label(schemePartList, filePartList);

      string schemeDeviceTagList = "Device tag list";
      string fileDeviceTagList = Path.Combine(outputPath, "Device tag list.txt");
      Label(schemeDeviceTagList, fileDeviceTagList);

      // Open folder in explorer
      Process.Start(outputPath);
    }

    private static void Label(string schemaName, string destinationFile)
    {
      if (File.Exists(destinationFile))
      {
        File.Delete(destinationFile);
      }

      ActionCallingContext actionCallingContext = new ActionCallingContext();
      actionCallingContext.AddParameter("configscheme", schemaName);
      actionCallingContext.AddParameter("destinationfile", destinationFile);
      actionCallingContext.AddParameter("language", "de_DE");
      new CommandLineInterpreter().Execute("label", actionCallingContext);
    }
  }
}