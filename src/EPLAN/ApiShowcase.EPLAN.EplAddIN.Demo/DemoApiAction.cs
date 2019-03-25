using System.Linq;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.MasterData;
using Eplan.EplApi.HEServices;

namespace ApiShowcase.EPLAN.EplAddIn.Demo
{
  // ReSharper disable once UnusedMember.Global
  class DemoApiAction : IEplAction
  {
    const string PROJECT_TEMPLATE_FILE_PATH =
      @"\\Mac\Home\Documents\GitHub\ibKastl.DemoData\2.7\Vorlagen\ibKastl_Basisprojekt_2.7_V01.zw9";

    const string PROJECT_LINK_FILE_PATH = @"C:\Test\EPLAN\TestProject.elk";

    const string WINDOW_MACRO_PATH =
      @"\\Mac\Home\Documents\GitHub\ibKastl.DemoData\2.7\Makros\ARTIKELMAKROS\ELEKTROTECHNIK\EINZELTEIL\MEW\MODULATOREN\ALLGEMEINE\127730.ems";

    public void GetActionProperties(ref ActionProperties actionProperties) {}

    public bool OnRegister(ref string name, ref int ordinal)
    {
      name = nameof(DemoApiAction);
      ordinal = 20;
      return true;
    }

    public bool Execute(ActionCallingContext actionCallingContext)
    {
      Project project = CreateProject();

      Page page = CreatePage(project);

      InsertMacros(project, page);

      OpenPageAndSyncInNavigator(page);

      return true;
    }

    private Project CreateProject()
    {
      var projectManager = new ProjectManager();

      Project project;

      // Open
      if (projectManager.ExistsProject(PROJECT_LINK_FILE_PATH))
      {
        project = projectManager.OpenProjects.FirstOrDefault(p => p.ProjectLinkFilePath.Equals(PROJECT_LINK_FILE_PATH));
        if (project == null)
        {
          project = projectManager.OpenProject(PROJECT_LINK_FILE_PATH, ProjectManager.OpenMode.Exclusive, true);
        }
        project.RemoveAllPages();
        return project;
      }

      // New
      project = projectManager.CreateProject(PROJECT_LINK_FILE_PATH, PROJECT_TEMPLATE_FILE_PATH);
      return project;

      // todo: Check path variables
      // todo: Check multi user conflicts
      // todo: Locking
    }

    private static Page CreatePage(Project project)
    {
      PagePropertyList pagePropertyList = new PagePropertyList();
      pagePropertyList.DESIGNATION_PLANT = "MYFUNCTION";
      pagePropertyList.DESIGNATION_LOCATION = "MYLOCATION";
      Page page = new Page(project, DocumentTypeManager.DocumentType.Circuit, pagePropertyList);
      return page;

      // todo: Locking
      // todo: Check if exists
    }

    private void InsertMacros(Project project, Page page)
    {
      PointD lastPoint;
      lastPoint = InsertMacro(WINDOW_MACRO_PATH, project, page, null);
      lastPoint = InsertMacro(WINDOW_MACRO_PATH, project, page, lastPoint);
      lastPoint = InsertMacro(WINDOW_MACRO_PATH, project, page, lastPoint);

      // todo: Path variables
      // todo: Check if exists
    }

    private PointD InsertMacro(string windowMacroPath, Project project, Page page, PointD? point)
    {
      SymbolMacro symbolMacro = new SymbolMacro();
      symbolMacro.Open(windowMacroPath, project);

      WindowMacro.Enums.RepresentationType representationType = symbolMacro.RepresentationTypes.First();
      int variant = 0;

      var offsetX = 40;
      var offsetY = 200;
      if (point == null)
      {
        point = new PointD(offsetX, offsetY);
      }
      else
      {
        point = new PointD(point.Value.X + offsetX, point.Value.Y);
      }

      var insert = new Insert();
      insert.SymbolMacro(symbolMacro,
                         representationType,
                         variant,
                         page,
                         point.Value,
                         Insert.MoveKind.Absolute,
                         WindowMacro.Enums.NumerationMode.Number);
      return page.GetBoundingBox().Last();

      // todo: Check variant
      // todo: Check RepresentationType
      // todo: Locking
    }

    public void OpenPageAndSyncInNavigator(Page page)
    {
      new Edit().OpenPageWithName(page.Project.ProjectLinkFilePath, page.IdentifyingName); // Open in GED
      new CommandLineInterpreter().Execute("XGedSelectPageAction"); // Select page
      new CommandLineInterpreter().Execute("XEsSyncPDDsAction"); // Sync selection
      new CommandLineInterpreter().Execute("XGedEscapeAction"); // Escape (page selection)
    }
  }
}