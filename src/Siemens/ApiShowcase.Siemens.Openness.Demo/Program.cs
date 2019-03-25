using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Siemens.Engineering;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;

namespace ApiShowcase.Siemens.Openness.Demo
{
  class Program
  {
    const string PROJECT_PATH = @"C:\Test\Siemens\TestProject\TestProject.ap15_1";

    private static TiaPortal _tiaPortal;
    private static PlcSoftware _plcSoftware;
    private static HmiTarget _hmiTarget;

    static void Main(string[] args)
    {
      // Start or connect to TIA Portal
      var tiaPortalProcesses = TiaPortal.GetProcesses();
      if (tiaPortalProcesses.Any())
      {
        _tiaPortal = tiaPortalProcesses.First().Attach();
      }
      else
      {
        _tiaPortal = new TiaPortal(TiaPortalMode.WithUserInterface);
      }

      try
      {
        var project = CreateProject();
        CreateDevices(project);

        LoadSclSources(_plcSoftware);

        InsertHmiPages(_hmiTarget);
      }
      finally
      {
        _tiaPortal.Dispose();
      }
    }

    private static void LoadSclSources(PlcSoftware plcSoftware)
    {
      // todo
    }

    private static void InsertHmiPages(HmiTarget hmiTarget)
    {
      // todo
    }

    private static void CreateDevices(Project project)
    {
      // Get existing
      GetDeviceSoftware(project);

      // Create
      if (_plcSoftware == null || _hmiTarget == null)
      {
        if (_plcSoftware == null)
        {
          project.Devices.CreateWithItem("OrderNumber:6ES7 510-1DJ01-0AB0/V2.0", "PLC_1", "MyPlc");  
        }
        if (_hmiTarget == null)
        {
          // todo
          project.Devices.CreateWithItem("OrderNumber:6AV2 123-2GA03-0AX0/15.1.0.0", "HMI_1", "MyHmi");  
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
          // Get software
          SoftwareContainer softwareContainer = ((IEngineeringServiceProvider)deviceItem).GetService<SoftwareContainer>();
          if (softwareContainer != null)
          {
            switch (softwareContainer.Software)
            {
              case HmiTarget hmiTarget:
                _hmiTarget = hmiTarget;
                break;
              case PlcSoftware plcSoftware:
                _plcSoftware = plcSoftware;
                break;
            }
          }

          // All found
          if (_hmiTarget != null && _plcSoftware != null)
          {
            return;
          }
        }
      }
    }

    private static Project CreateProject()
    {
      Console.WriteLine("Create project: " + PROJECT_PATH);
      Project project;

      // Open
      if (File.Exists(PROJECT_PATH))
      {
        FileInfo fileInfo = new FileInfo(PROJECT_PATH);
        project = _tiaPortal.Projects.FirstOrDefault(obj => obj.Path.FullName.Equals(PROJECT_PATH));
        if (project == null)
        {
          project = _tiaPortal.Projects.Open(fileInfo);
        }
      }

      // Create
      else
      {
        string projectName = Path.GetFileNameWithoutExtension(PROJECT_PATH);
        var directory = Path.GetDirectoryName(PROJECT_PATH);
        directory = Path.GetDirectoryName(directory); // Project folder name is same as project name
        DirectoryInfo directoryInfo = new DirectoryInfo(directory ?? throw new InvalidOperationException());
        project = _tiaPortal.Projects.Create(directoryInfo, projectName);
      }
      return project;
    }
  }
}
