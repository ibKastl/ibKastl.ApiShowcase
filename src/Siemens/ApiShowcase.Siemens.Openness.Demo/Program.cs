using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Principal;
using Microsoft.Win32;
using Siemens.Engineering;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.ExternalSources;

namespace ApiShowcase.Siemens.Openness.Demo
{
  class Program
  {
    const string PROJECT_PATH = @"C:\Test\Siemens\TestProject\TestProject.ap15_1";

    private static TiaPortal _tiaPortal;
    private static PlcSoftware _plcSoftware;
    private static HmiTarget _hmiTarget;
    private static readonly List<Node> Nodes = new List<Node>();

    [STAThread]
    static void Main()
    {
      Stopwatch sw = new Stopwatch();
      sw.Restart();

      InitTiaPortal();

      try
      {
        var project = CreateProjectIfNotExists(PROJECT_PATH);
        CreateDevices(project);
        CreateSubnet(project);
        LoadSclSources(_plcSoftware);
        InsertHmiPages(_hmiTarget);
        Compile();
        ShowBlockInEditor();
        project.Save();
      }
      finally
      {
        TiaPortalProcess tiaProcess = _tiaPortal.GetCurrentProcess();
        var process = Process.GetProcessById(tiaProcess.Id);
        _tiaPortal.Dispose();
        Console.WriteLine($"Finished in {sw.Elapsed.Minutes}:{sw.Elapsed.Seconds}");
        WindowHelper.BringProcessToFront(process);

        Console.WriteLine();
        Console.WriteLine("Press any key to continue...");
        Console.ReadKey();
      }
    }

    private static void CreateSubnet(Project project)
    {
      // Clear
      foreach (var projectSubnet in project.Subnets.ToList())
      {
        projectSubnet.Delete();
      }

      // Connect
      Subnet subnet = project.Subnets.Create("System:Subnet.Ethernet", "MySubnet");
      foreach (var node in Nodes)
      {
        node.ConnectToSubnet(subnet);
      }
    }

    private static void InitTiaPortal()
    {
      Console.WriteLine("Check firewall...");

      SetTiaPortalFirewall();

      // Open TIA
      Console.WriteLine("Init TIA Portal...");
      var tiaPortalProcesses = TiaPortal.GetProcesses();
      if (tiaPortalProcesses.Any())
      {
        _tiaPortal = tiaPortalProcesses.First().Attach();
      }
      else
      {
        _tiaPortal = new TiaPortal(TiaPortalMode.WithUserInterface);
      }
    }

    public static void SetTiaPortalFirewall()
    {
      // Check if admin
      WindowsIdentity identity = WindowsIdentity.GetCurrent();
      WindowsPrincipal principal = new WindowsPrincipal(identity);
      bool isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
      if (!isAdmin)
      {
        return;
      }

      Assembly assembly = Assembly.GetExecutingAssembly();
      string exePath = assembly.Location;

      // Get hash
      HashAlgorithm hashAlgorithm = SHA256.Create();
      FileStream stream = File.OpenRead(exePath);
      byte[] hash = hashAlgorithm.ComputeHash(stream);
      string convertedHash = Convert.ToBase64String(hash);

      // Get date
      FileInfo fileInfo = new FileInfo(exePath);
      DateTime lastWriteTimeUtc = fileInfo.LastWriteTimeUtc;
      string lastWriteTimeUtcFormatted = lastWriteTimeUtc.ToString("yyyy'/'MM'/'dd HH:mm:ss.fff");

      // Get execution version
      AssemblyName siemensAssembly = Assembly.GetExecutingAssembly().GetReferencedAssemblies()
                                             .First(obj => obj.Name.Equals("Siemens.Engineering"));
      string version = siemensAssembly.Version.ToString(2);

      // Set key and values
      string keyFullName = $@"SOFTWARE\Siemens\Automation\Openness\{version}\Whitelist\{fileInfo.Name}\Entry";
      RegistryKey key = Registry.LocalMachine.CreateSubKey(keyFullName);
      if (key == null)
      {
        throw new Exception("Key note found: " + keyFullName);
      }
      key.SetValue("Path", exePath);
      key.SetValue("DateModified", lastWriteTimeUtcFormatted);
      key.SetValue("FileHash", convertedHash);
    }

    public static Project CreateProjectIfNotExists(string projectFullPath)
    {
      Console.WriteLine("Create project...");
      Project project;

      // Open
      if (File.Exists(projectFullPath))
      {
        FileInfo fileInfo = new FileInfo(projectFullPath);
        project = _tiaPortal.Projects.FirstOrDefault(obj => obj.Path.FullName.Equals(projectFullPath));
        if (project == null)
        {
          project = _tiaPortal.Projects.Open(fileInfo);
        }
      }

      // Create
      else
      {
        string projectName = Path.GetFileNameWithoutExtension(projectFullPath);
        var directory = Path.GetDirectoryName(projectFullPath);
        directory = Path.GetDirectoryName(directory); // Project folder name is same as project name
        DirectoryInfo directoryInfo = new DirectoryInfo(directory ?? throw new InvalidOperationException());
        project = _tiaPortal.Projects.Create(directoryInfo, projectName);
      }
      return project;
    }

    private static void CreateDevices(Project project)
    {
      Console.WriteLine("Get devices...");

      // Get existing
      GetDeviceSoftware(project);

      // Create
      if (_plcSoftware == null || _hmiTarget == null)
      {
        if (_plcSoftware == null)
        {
          Console.WriteLine("Create PLC...");
          project.Devices.CreateWithItem("OrderNumber:6ES7 211-1AD30-0XB0/V2.2", "PLC_1", null);
        }
        if (_hmiTarget == null)
        {
          Console.WriteLine("Create HMI...");
          project.Devices.CreateWithItem("OrderNumber:6AV2 123-2GB03-0AX0/15.1.0.0", "HMI_1",
                                         null); // Name not allowed
        }

        GetDeviceSoftware(project);
      }
    }

    private static void GetDeviceSoftware(Project project)
    {
      foreach (var device in project.Devices)
      {
        foreach (var deviceItem in device.DeviceItems)
        {
          GetDeviceSoftware(deviceItem);
        }
      }
    }

    private static void GetDeviceSoftware(DeviceItem deviceItem)
    {
      // Get software
      SoftwareContainer softwareContainer =
        ((IEngineeringServiceProvider)deviceItem).GetService<SoftwareContainer>();
      if (softwareContainer != null)
      {
        switch (softwareContainer.Software)
        {
          case HmiTarget hmiTarget:
            Console.WriteLine("Get HmiTarget...");
            _hmiTarget = hmiTarget;
            break;
          case PlcSoftware plcSoftware:
            Console.WriteLine("Get PlcSoftware...");
            _plcSoftware = plcSoftware;
            break;
        }
      }

      // Get network interface
      var networkInterface = deviceItem.GetService<NetworkInterface>();
      if (networkInterface != null)
      {
        foreach (var node in networkInterface.Nodes)
        {
          Nodes.Add(node);
        }
      }

      foreach (var subDeviceItem in deviceItem.DeviceItems)
      {
        GetDeviceSoftware(subDeviceItem);
      }
    }

    private static void LoadSclSources(PlcSoftware plcSoftware)
    {
      // Create source
      var sourceName = "CheckArray";
      var source = plcSoftware.ExternalSourceGroup.ExternalSources.FirstOrDefault(obj => obj.Name.Equals(sourceName));
      if (source == null)
      {
        Console.WriteLine("Load SCL sources...");
        string sourcePath = @"\\Mac\Home\Documents\GitHub\ibKastl.ApiShowcase\data\Siemens\CheckArray.scl";
        source = plcSoftware.ExternalSourceGroup.ExternalSources.CreateFromFile(sourceName, sourcePath);
      }

      // Create block from source
      PlcBlock block = plcSoftware.BlockGroup.Blocks.FirstOrDefault(obj => obj.Name.Equals(sourceName));
      if (block == null)
      {
        Console.WriteLine("Generate block...");
        GenerateBlockOption options = GenerateBlockOption.KeepOnError;
        source.GenerateBlocksFromSource(options);
      }
    }

    private static void InsertHmiPages(HmiTarget hmiTarget)
    {
      Console.WriteLine("Insert screens from file...");

      string screenFile = @"\\Mac\Home\Documents\GitHub\ibKastl.ApiShowcase\data\Siemens\Screen.xml";
      FileInfo fileInfo = new FileInfo(screenFile);

      // Export first time
      if (!fileInfo.Exists)
      {
        hmiTarget.ScreenFolder.Screens.Last().Export(fileInfo, ExportOptions.WithDefaults);
      }

      hmiTarget.ScreenFolder.Screens.Import(fileInfo, ImportOptions.Override);
    }

    private static void Compile()
    {
      Console.WriteLine("Compile...");
      _hmiTarget.GetService<ICompilable>().Compile();
      _plcSoftware.GetService<ICompilable>().Compile();
    }

    private static void ShowBlockInEditor()
    {
      Console.WriteLine("Show block in editor...");
      var block = _plcSoftware.BlockGroup.Blocks
                              .OfType<OB>()
                              .First(obj => obj.Number.Equals(1));
      block.ShowInEditor();
    }
  }
}