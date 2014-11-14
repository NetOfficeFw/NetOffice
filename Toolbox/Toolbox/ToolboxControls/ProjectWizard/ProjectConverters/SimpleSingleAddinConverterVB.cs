using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.ProjectConverters
{
    internal class SimpleSingleAddinConverterVB : Converter
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

        internal SimpleSingleAddinConverterVB(ProjectOptions options) : base(options)
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

            _addinFile = _addinFile.Replace("$safeprojectname$", Options.AssemblyName);
            _addinFile = _addinFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingItems());
            _addinFile = _addinFile.Replace("$randomGuid$", Guid.NewGuid().ToString().ToUpper());
            _addinFile = _addinFile.Replace("$applicationField$", String.Format("\tPrivate _application As {0}.Application", Options.OfficeApps[0]));
            _addinFile = _addinFile.Replace("$register$", ReadProjectTemplateFile("SimpleSingleAddinVB.Register.txt").Replace("%HiveKey%", Options.HiveKey).Replace("%Name%", Options.AssemblyName).Replace("%Description%", Options.AssemblyDescription).Replace("%LoadBehavior%", Options.LoadBehaviour.ToString()).Replace("%officAddinKey%", Options.RegistryKeys[0] + "\\" + Options.AssemblyName + ".Addin"));
            _addinFile = _addinFile.Replace("$unregister$", ReadProjectTemplateFile("SimpleSingleAddinVB.UnRegister.txt").Replace("%HiveKey%", Options.HiveKey).Replace("%officAddinKey%", Options.RegistryKeys[0] + "\\" + Options.AssemblyName + ".Addin"));
            _addinFile = _addinFile.Replace("$applicationConstruction$", String.Format("\t\t\t_application = new {0}.Application(Nothing, application)", Options.OfficeApps[0]));
            _addinFile = _addinFile.Replace("$applicationDestroy$", "\t\t\tIf(Not IsNothing(_application))\r\n\t\t\t\t_application.Dispose()\r\n\t\t\tEnd If");

            _assemblyFile = _assemblyFile.Replace("$safeprojectname$", Options.AssemblyName);
            _assemblyFile = _assemblyFile.Replace("$safeprojectdescription$", Options.AssemblyDescription);
            _assemblyFile = _assemblyFile.Replace("$assemblyguid$", Guid.NewGuid().ToString().ToUpper());

            _taskPaneFile = _taskPaneFile.Replace("$safeprojectname$", Options.AssemblyName);
            _taskPaneFile = _taskPaneFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingItems());

            _taskPaneDesignerFile = _taskPaneDesignerFile.Replace("$safeprojectname$", Options.AssemblyName);
            _taskPaneDesignerFile = _taskPaneDesignerFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingItems());

            if (Options.UseRibbonUI)
            {
                _addinFile = _addinFile.Replace("$ribbonDefine$", " , Office.IRibbonExtensibility");
                _addinFile = _addinFile.Replace("$ribbonImplement$", ReadProjectTemplateFile("SimpleSingleAddinVB.RibbonImplement.txt"));
                _projectFile = _projectFile.Replace("$ribbonFileReference$", "  <ItemGroup>\r\n    <EmbeddedResource Include=\"RibbonUI.xml\" />\r\n  </ItemGroup>");
                _ribbonFile = _ribbonFile.Replace("$safeprojectname$", Options.AssemblyName);
            }
            else
            {
                _addinFile = _addinFile.Replace("$ribbonDefine$", String.Empty);
                _addinFile = _addinFile.Replace("$ribbonImplement$", String.Empty);
                _projectFile = _projectFile.Replace("$ribbonFileReference$", String.Empty);
            }

            if (Options.UseTaskPane)
            {
                _addinFile = _addinFile.Replace("$taskpaneDefine$", " , Office.ICustomTaskPaneConsumer");
                _addinFile = _addinFile.Replace("$taskpaneImplement$", ReadProjectTemplateFile("SimpleSingleAddinVB.TaskPaneImplement.txt").Replace("$safeprojectname$", Options.AssemblyName));
                _addinFile = _addinFile.Replace("$taskpaneField$", "\tPrivate _mytaskPane As MyTaskPane");
            }
            else
            {
                _addinFile = _addinFile.Replace("$taskpaneDefine$", String.Empty);
                _addinFile = _addinFile.Replace("$taskpaneImplement$", String.Empty);
                _addinFile = _addinFile.Replace("$taskpaneField$", String.Empty);
            }

            if (Options.UseClassicUI)
            {
                _addinFile = _addinFile.Replace("$classicUICreateCall$", "\t\t\tCreateUserInterface()");
                _addinFile = _addinFile.Replace("$classicUIRemoveCall$", "\t\t\tRemoveUserInterface()");

                string template = "#Region \"Classic UI Methods\"\r\n\r\n" +
                                  "\tPrivate Sub CreateUserInterface()\r\n\t\t\r\n            \r\n\tEnd Sub\r\n\r\n" +
                                  "\tPrivate Sub RemoveUserInterface()\r\n\t\t\r\n            \r\n\tEnd Sub\r\n\r\n" +
                                  "#End Region";

                _addinFile = _addinFile.Replace("$classicUICreateRemoveMethod$", template);
            }
            else
            {
                _addinFile = _addinFile.Replace("$classicUICreateCall$", String.Empty);
                _addinFile = _addinFile.Replace("$classicUIRemoveCall$", String.Empty);
                _addinFile = _addinFile.Replace("$classicUICreateRemoveMethod$", String.Empty);
            }

            if (Options.UseRibbonUI)
                _addinFile = _addinFile.Replace("$readRessource$", ReadProjectTemplateFile("SimpleSingleAddinVB.ReadResource.txt"));
            else
                _addinFile = _addinFile.Replace("$readRessource$", String.Empty);

            _addinFile = ValidateFileContentFormat(_addinFile);
        }

        private void ReadRessourceFiles()
        {
            _taskPaneFile = ReadProjectTemplateFile("SimpleSingleAddinVB.TaskPane.txt");
            _taskPaneDesignerFile = ReadProjectTemplateFile("SimpleSingleAddinVB.TaskPane_Designer.txt");
            _ribbonFile = ReadProjectTemplateFile("SimpleSingleAddinVB.RibbonUI.txt");
            _appDesignerFile = ReadProjectTemplateFile("SimpleSingleAddinVB.Application_Designer.txt");
            _myApplicationFile = ReadProjectTemplateFile("SimpleSingleAddinVB.Application_myapp.txt");
            _ressourceDesgnerFile = ReadProjectTemplateFile("SimpleSingleAddinVB.Resources_Designer.txt");
            _ressourceResFile = ReadProjectTemplateFile("SimpleSingleAddinVB.Resources_resx.txt");
            _settingDesignerFile = ReadProjectTemplateFile("SimpleSingleAddinVB.Settings_Designer.txt");
            _settingsSettingsFile = ReadProjectTemplateFile("SimpleSingleAddinVB.Settings_settings.txt");
            _solutionFile = ReadProjectTemplateFile("SimpleSingleAddinVB.Solution.txt");
            _projectFile = ReadProjectTemplateFile("SimpleSingleAddinVB.Project.txt");
            _projectUserFile = ReadProjectTemplateFile("SimpleSingleAddinVB.Project_User.txt");
            _addinFile = ReadProjectTemplateFile("SimpleSingleAddinVB.Addin.txt");
            _assemblyFile = ReadProjectTemplateFile("SimpleSingleAddinVB.AssemblyInfo.txt");
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
