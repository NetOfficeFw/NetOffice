using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.ProjectConverters
{
    internal class ToolsMultiAddinConverterVB : Converter
    {
        #region Fields

        private string _taskPaneFile;
        private string _taskPaneDesignerFile;
        private string _ribbonFile;
        private string _appDesignerFile;
        private string _myApplicationFile;
        private string _ressourceDesgnerFile;
        private string _ressourceResFile;
        private string _settingDesignerFile;
        private string _settingsSettingsFile;
        private string _solutionFile;
        private string _projectFile;
        private string _projectUserFile;
        private string _addinFile;
        private string _assemblyFile;
        private Guid   _projectGuid;
        private bool   _applicationFound;

        #endregion

        #region Ctor

        internal ToolsMultiAddinConverterVB(ProjectOptions options) : base(options)
        { 
        
        }

        #endregion

        #region Overrides

        public override string CreateSolution()
        {
            ReadRessourceFiles();
            ReplaceMarker();
            WriteResultFilesToTempFolder();
            CopyUsedNetOfficeAssembliesToTempTarget();
            MoveTempSolutionFolderToTarget();
            return TargetSolutionFile;
        }

        #endregion

        #region Methods

        private void ReplaceMarker()
        {
            _projectGuid = Guid.NewGuid();

            _ressourceDesgnerFile = _ressourceDesgnerFile.Replace("$safeprojectname$", Options.AssemblyName);

            _settingDesignerFile = _settingDesignerFile.Replace("$safeprojectname$", Options.AssemblyName);

            _solutionFile = _solutionFile.Replace("$safeprojectname$", Options.AssemblyName);
            _solutionFile = _solutionFile.Replace("$projectguid$", _projectGuid.ToString().ToUpper());
            _solutionFile = _solutionFile.Replace("$solutionformat$", this.SolutionFormats[Options.IDE]);
            _solutionFile = _solutionFile.Replace("$ideversion$", this.Environments[Options.IDE, Options.Language]);

            _projectFile = _projectFile.Replace("$safeprojectname$", Options.AssemblyName);
            _projectFile = _projectFile.Replace("$projectguid$", _projectGuid.ToString().ToUpper());
            _projectFile = _projectFile.Replace("$toolsversion$", this.Tools[Options.IDE]);
            _projectFile = _projectFile.Replace("$targetframeworkversion$", this.Runtimes[Options.NetRuntime]);
            _projectFile = _projectFile.Replace("$assemblyReferences$", this.GetNetOfficeProjectReferenceItems());

            string applicationClassID = TryGetRegistryValue(RegistryHive.HKEY_Local_Machine, String.Format("Software\\Classes\\{0}.Application\\CLSID", Options.OfficeApps[0])) as string;
            if (!String.IsNullOrWhiteSpace(applicationClassID))
            {
                string applicationPath = TryGetRegistryValue(RegistryHive.HKEY_Local_Machine, String.Format("Software\\Classes\\CLSID\\{0}\\LocalServer32", applicationClassID)) as string;
                if (!String.IsNullOrWhiteSpace(applicationPath))
                {
                    int argumentStartIndex = applicationPath.IndexOf("/", 0);
                    if (argumentStartIndex > -1)
                        applicationPath = applicationPath.Substring(0, argumentStartIndex).Trim();
                    _projectUserFile = _projectUserFile.Replace("$toolsversion$", this.Tools[Options.IDE]);
                    _projectUserFile = _projectUserFile.Replace("$path$", applicationPath);
                    _applicationFound = true;
                }
            }

            if (Options.UseTaskPane)
                _projectFile = _projectFile.Replace("$taskpaneFileReference$", "  <Compile Include=\"MyTaskPane.vb\">\r\n   <SubType>UserControl</SubType>\r\n   </Compile>\r\n  <Compile Include=\"MyTaskPane.Designer.vb\">\r\n    <DependentUpon>MyTaskPane.vb</DependentUpon>\r\n  </Compile>");
            else
                _projectFile = _projectFile.Replace("$taskpaneFileReference$", String.Empty);
            
            _addinFile = _addinFile.Replace("$appName$", Options.OfficeApps[0]);
            _addinFile = _addinFile.Replace("$safeprojectname$", Options.AssemblyName);
            _addinFile = _addinFile.Replace("$description$", Options.AssemblyDescription);
            _addinFile = _addinFile.Replace("$loadbeahviour$", ConvertLoadBehavoir(Options.LoadBehaviour));

            _addinFile = _addinFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingToolsItems());
            _addinFile = _addinFile.Replace("$randomGuid$", Guid.NewGuid().ToString().ToUpper());
            _addinFile = _addinFile.Replace("$applicationConstruction$", String.Format("\t\t\t_application = new {0}.Application(Nothing, application)", Options.OfficeApps[0]));
            _addinFile = _addinFile.Replace("$applicationDestroy$", "\t\t\tIf(Not IsNothing(_application))\r\n\t\t\t\t_application.Dispose()\r\n\t\t\tEnd If");

            _assemblyFile = _assemblyFile.Replace("$safeprojectname$", Options.AssemblyName);
            _assemblyFile = _assemblyFile.Replace("$safeprojectdescription$", Options.AssemblyDescription);
            _assemblyFile = _assemblyFile.Replace("$assemblyguid$", Guid.NewGuid().ToString().ToUpper());

            _taskPaneFile = _taskPaneFile.Replace("$safeprojectname$", Options.AssemblyName);
            _taskPaneFile = _taskPaneFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingToolsItems());
            _taskPaneFile = _taskPaneFile.Replace("$appName$", Options.OfficeApps[0]);

            _taskPaneDesignerFile = _taskPaneDesignerFile.Replace("$safeprojectname$", Options.AssemblyName);
            _taskPaneDesignerFile = _taskPaneDesignerFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingToolsItems());

            string getVersion = "\t\tConsole.WriteLine(\"Addin started in {0}\", Application.InstanceFriendlyName)";
            _addinFile = _addinFile.Replace("$getversion$", getVersion);

            string attribute2String = "<MultiRegister(";
            foreach (var item in Options.OfficeApps)
                attribute2String += "RegisterIn." + item + ", ";
            attribute2String = attribute2String.Substring(0, attribute2String.Length - 2);
            attribute2String += ")>";
            _addinFile = _addinFile.Replace("$attributes2$", attribute2String);

            string attributeString = "";
            if (Options.HiveKey == "LocalMachine")
            {
                attributeString = "RegistryLocation(RegistrySaveLocation.LocalMachine)";
            }
            else
            {
                attributeString = "RegistryLocation(RegistrySaveLocation.InstallScopeCurrentUser)";
            }

            if (Options.UseRibbonUI)
            {
                attributeString += ", CustomUI(\"RibbonUI.xml\", True)".Replace("$safeprojectname$", Options.AssemblyName);
                _addinFile = _addinFile.Replace("$ribbonProperty$", "\tFriend Property RibbonUI() As Office.IRibbonUI\r\n"
                                                                    + "\t\tGet\r\n"
                                                                    + "\t\t\tReturn _ribbonUI\r\n"
                                                                    + "\t\tEnd Get\r\n"
                                                                    + "\t\tPrivate Set(ByVal Value As Office.IRibbonUI)\r\n"
                                                                    + "\t\t\t_ribbonUI = Value\r\n"
                                                                    + "\t\tEnd Set\r\n"
                                                                    + "\tEnd Property\r\n"
                                                                    + "\tPrivate _ribbonUI As Office.IRibbonUI");
                _addinFile = _addinFile.Replace("$ribbonImplement$", ReadProjectTemplateFile("ToolsSingleAddinVB.RibbonImplement.txt"));
                _projectFile = _projectFile.Replace("$ribbonFileReference$", "  <ItemGroup>\r\n    <EmbeddedResource Include=\"RibbonUI.xml\" />\r\n  </ItemGroup>");
                _ribbonFile = _ribbonFile.Replace("$safeprojectname$", Options.AssemblyName);
                _addinFile = _addinFile.Replace("$ribbonLoad$", ReadProjectTemplateFile("ToolsSingleAddinVB.RibbonImplement.txt"));
                _addinFile = _addinFile.Replace("$safeprojectname$", Options.AssemblyName);
            }
            else
            {
                _addinFile = _addinFile.Replace("$ribbonProperty$", String.Empty);
                _addinFile = _addinFile.Replace("$ribbonDefine$", String.Empty);
                _addinFile = _addinFile.Replace("$ribbonImplement$", String.Empty);
                _projectFile = _projectFile.Replace("$ribbonFileReference$", String.Empty);
                _addinFile = _addinFile.Replace("$ribbonLoad$", String.Empty);
            }

            if (Options.UseTaskPane)
            {

                if (Options.UseRibbonUI || Options.HiveKey == "LocalMachine")
                    attributeString += ", CustomPane(GetType(MyTaskPane), \"My TaskPane\", true, PaneDockPosition.msoCTPDockPositionRight)";
                else
                    attributeString += "CustomPane(GetType(MyTaskPane), \"My TaskPane\", true, PaneDockPosition.msoCTPDockPositionRight)";

                _addinFile = _addinFile.Replace("$taskpaneDefine$", " , Office.ICustomTaskPaneConsumer");
                _addinFile = _addinFile.Replace("$taskpaneField$", "\tPrivate _mytaskPane As MyTaskPane");
            }
            else
            {
                _addinFile = _addinFile.Replace("$taskpaneDefine$", String.Empty);
                _addinFile = _addinFile.Replace("$taskpaneField$", String.Empty);
            }

            if (Options.UseClassicUI)
            {
                _addinFile = _addinFile.Replace("$classicUICreateCall$", "\t\tCreateUserInterface()");
                _addinFile = _addinFile.Replace("$classicUIRemoveCall$", "\t\tRemoveUserInterface()");

                string template = 
                                  "\tPrivate Sub CreateUserInterface()\r\n\r\n\tEnd Sub\r\n\r\n" +
                                  "\tPrivate Sub RemoveUserInterface()\r\n\r\n\tEnd Sub";

                _addinFile = _addinFile.Replace("$classicUICreateRemoveMethod$", template);
            }
            else
            {
                _addinFile = _addinFile.Replace("$classicUICreateCall$", String.Empty);
                _addinFile = _addinFile.Replace("$classicUIRemoveCall$", String.Empty);
                _addinFile = _addinFile.Replace("$classicUICreateRemoveMethod$", String.Empty);
            }

            if (Options.UseToogle)
            {
                _addinFile = _addinFile.Replace("$tooglecode$", ReadProjectTemplateFile("ToolsSingleAddinVB.RibbonToogleImplement.txt"));
                _taskPaneFile = _taskPaneFile.Replace("$tooglecall$", "\t\tIf Not IsNothing(ParentAddin) And Not IsNothing(ParentAddin.RibbonUI) Then\r\n\t\t\tParentAddin.RibbonUI.InvalidateControl(\"tooglePaneVisibleButton\")\r\n\t\tEnd If\r\n");
                _ribbonFile = _ribbonFile.Replace("$tooglebutton$", "           <toggleButton id=\"tooglePaneVisibleButton\" label=\"TaskPane\" imageMso=\"CreateFormBlankForm\" size=\"large\" getPressed=\"TooglePaneVisibleButton_GetPressed\" onAction=\"TooglePaneVisibleButton_Click\" />");
            }
            else
            {
                _addinFile = _addinFile.Replace("$tooglecode$", String.Empty);
                _taskPaneFile = _taskPaneFile.Replace("\r\n$tooglecall$", String.Empty);
                _ribbonFile = _ribbonFile.Replace("\r\n$tooglebutton$", String.Empty);
            }

            if (!String.IsNullOrWhiteSpace(attributeString))
                attributeString = "<" + attributeString + ">";
            _addinFile = _addinFile.Replace("$attributes$", attributeString);

            _addinFile = ValidateFileContentFormat(_addinFile);
        }

        private void ReadRessourceFiles()
        {
            _taskPaneFile = ReadProjectTemplateFile("ToolsMultiAddinVB.TaskPane.txt");
            _taskPaneDesignerFile = ReadProjectTemplateFile("ToolsMultiAddinVB.TaskPane_Designer.txt");
            _ribbonFile = ReadProjectTemplateFile("ToolsMultiAddinVB.RibbonUI.txt");
            _appDesignerFile = ReadProjectTemplateFile("ToolsMultiAddinVB.Application_Designer.txt");
            _myApplicationFile = ReadProjectTemplateFile("ToolsMultiAddinVB.Application_myapp.txt");
            _ressourceDesgnerFile = ReadProjectTemplateFile("ToolsMultiAddinVB.Resources_Designer.txt");
            _ressourceResFile = ReadProjectTemplateFile("ToolsMultiAddinVB.Resources_resx.txt");
            _settingDesignerFile = ReadProjectTemplateFile("ToolsMultiAddinVB.Settings_Designer.txt");
            _settingsSettingsFile = ReadProjectTemplateFile("ToolsMultiAddinVB.Settings_settings.txt");
            _solutionFile = ReadProjectTemplateFile("ToolsMultiAddinVB.Solution.txt");
            _projectFile = ReadProjectTemplateFile("ToolsMultiAddinVB.Project.txt");
            _projectUserFile = ReadProjectTemplateFile("ToolsMultiAddinVB.Project_User.txt");
            _addinFile = ReadProjectTemplateFile("ToolsMultiAddinVB.Addin.txt");
            _assemblyFile = ReadProjectTemplateFile("ToolsMultiAddinVB.AssemblyInfo.txt");
        }

        private void WriteResultFilesToTempFolder()
        {
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Application.Designer.vb"), _appDesignerFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Application.myapp"), _myApplicationFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Resources.Designer.vb"), _ressourceDesgnerFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Resources.resx"), _ressourceResFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Settings.Designer.vb"), _settingDesignerFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Settings.settings"), _settingsSettingsFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempSolutionPath, String.Format("{0}.sln", Options.AssemblyName)), _solutionFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempProjectPath, String.Format("{0}.vbproj", Options.AssemblyName)), _projectFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempProjectPath, "Addin.vb"), _addinFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "AssemblyInfo.vb"), _assemblyFile, Encoding.UTF8);

            if (_applicationFound)
                File.AppendAllText(Path.Combine(TempProjectPath, String.Format("{0}.vbproj.user", Options.AssemblyName)), _projectUserFile, Encoding.UTF8);

            if (Options.UseRibbonUI)
                File.AppendAllText(Path.Combine(TempProjectPath, "RibbonUI.xml"), _ribbonFile, Encoding.UTF8);
            if (Options.UseTaskPane)
            {
                File.AppendAllText(Path.Combine(TempProjectPath, "MyTaskPane.vb"), _taskPaneFile, Encoding.UTF8);
                File.AppendAllText(Path.Combine(TempProjectPath, "MyTaskPane.Designer.vb"), _taskPaneDesignerFile, Encoding.UTF8);
            }
        }

        #endregion
    }
}
