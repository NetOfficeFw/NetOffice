using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.ProjectConverters
{
    internal class SimpleSingleAddinConverterCS : Converter
    {
        #region Fields

        private string _taskPaneFile;
        private string _taskPaneDesignerFile;
        private string _ribbonFile;
        private string _solutionFile;
        private string _projectFile;
        private string _projectUserFile;
        private string _addinFile;
        private string _assemblyFile;
        private Guid   _projectGuid;
        private bool   _applicationFound;

        #endregion

        #region Ctor

        internal SimpleSingleAddinConverterCS(ProjectOptions options) : base(options)
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
                _projectFile = _projectFile.Replace("$taskpaneFileReference$", "  <Compile Include=\"MyTaskPane.cs\">\r\n   <SubType>UserControl</SubType>\r\n   </Compile>\r\n  <Compile Include=\"MyTaskPane.Designer.cs\">\r\n    <DependentUpon>MyTaskPane.cs</DependentUpon>\r\n  </Compile>");
            else
                _projectFile = _projectFile.Replace("$taskpaneFileReference$", String.Empty);

            _addinFile = _addinFile.Replace("$safeprojectname$", Options.AssemblyName);
            _addinFile = _addinFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingItems());
            _addinFile = _addinFile.Replace("$randomGuid$", Guid.NewGuid().ToString().ToUpper());
            _addinFile = _addinFile.Replace("$applicationField$", String.Format("\t\tprivate {0}.Application _application;", Options.OfficeApps[0]));
            _addinFile = _addinFile.Replace("$register$", ReadProjectTemplateFile("SimpleSingleAddinCS.Register.txt").Replace("%HiveKey%", Options.HiveKey).Replace("%Name%", Options.AssemblyName).Replace("%Description%", Options.AssemblyDescription).Replace("%LoadBehavior%", Options.LoadBehaviour.ToString()).Replace("%officAddinKey%", Options.RegistryKeys[0] + "\\" + Options.AssemblyName + ".Addin"));
            _addinFile = _addinFile.Replace("$unregister$", ReadProjectTemplateFile("SimpleSingleAddinCS.UnRegister.txt").Replace("%HiveKey%", Options.HiveKey).Replace("%officAddinKey%", Options.RegistryKeys[0] + "\\" + Options.AssemblyName + ".Addin"));
            _addinFile = _addinFile.Replace("$applicationConstruction$", String.Format("\t\t\t\t_application = new {0}.Application(null, application);", Options.OfficeApps[0]));
            _addinFile = _addinFile.Replace("$applicationDestroy$", "\t\t\t\tif(null != _application)\r\n\t\t\t\t\t_application.Dispose();");

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
                _addinFile = _addinFile.Replace("$ribbonImplement$", ReadProjectTemplateFile("SimpleSingleAddinCS.RibbonImplement.txt"));
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
                _addinFile = _addinFile.Replace("$taskpaneImplement$", ReadProjectTemplateFile("SimpleSingleAddinCS.TaskPaneImplement.txt").Replace("$safeprojectname$", Options.AssemblyName));
                _addinFile = _addinFile.Replace("$taskpaneField$", "\t\tprivate MyTaskPane _mytaskPane;");
                
            }
            else
            {
                _addinFile = _addinFile.Replace("$taskpaneDefine$", String.Empty);
                _addinFile = _addinFile.Replace("$taskpaneImplement$", String.Empty);
                _addinFile = _addinFile.Replace("$taskpaneField$", String.Empty);
            }

            if (Options.UseClassicUI)
            {
                _addinFile = _addinFile.Replace("$classicUICreateCall$", "\t\t\t\tCreateUserInterface();");
                _addinFile = _addinFile.Replace("$classicUIRemoveCall$", "\t\t\t\tRemoveUserInterface();");

                string template = "\t\t#region Classic UI Methods\r\n\r\n" +
                                  "\t\tprivate void CreateUserInterface()\r\n\t\t{\r\n            \r\n\t\t}\r\n\r\n" +
                                  "\t\tprivate void RemoveUserInterface()\r\n\t\t{\r\n            \r\n\t\t}\r\n\r\n" +
                                  "\t\t#endregion";

                _addinFile = _addinFile.Replace("$classicUICreateRemoveMethod$", template);
            }
            else
            {
                _addinFile = _addinFile.Replace("$classicUICreateCall$", String.Empty);
                _addinFile = _addinFile.Replace("$classicUIRemoveCall$", String.Empty);
                _addinFile = _addinFile.Replace("$classicUICreateRemoveMethod$", String.Empty);
            }

            if (Options.UseRibbonUI)
                _addinFile = _addinFile.Replace("$readRessource$", ReadProjectTemplateFile("SimpleSingleAddinCS.ReadResource.txt"));
            else 
                _addinFile = _addinFile.Replace("$readRessource$", String.Empty);

            _addinFile = ValidateFileContentFormat(_addinFile);
        }

        private void ReadRessourceFiles()
        {
            _taskPaneFile = ReadProjectTemplateFile("SimpleSingleAddinCS.TaskPane.txt");
            _taskPaneDesignerFile = ReadProjectTemplateFile("SimpleSingleAddinCS.TaskPane_Designer.txt");
            _ribbonFile = ReadProjectTemplateFile("SimpleSingleAddinCS.RibbonUI.txt");
            _solutionFile = ReadProjectTemplateFile("SimpleSingleAddinCS.Solution.txt");
            _projectFile = ReadProjectTemplateFile("SimpleSingleAddinCS.Project.txt");
            _projectUserFile = ReadProjectTemplateFile("SimpleSingleAddinCS.Project_User.txt");
            _addinFile = ReadProjectTemplateFile("SimpleSingleAddinCS.Addin.txt");
            _assemblyFile = ReadProjectTemplateFile("SimpleSingleAddinCS.AssemblyInfo.txt");
        }

        private void WriteResultFilesToTempFolder()
        {            
            File.AppendAllText(Path.Combine(TempSolutionPath, String.Format("{0}.sln", Options.AssemblyName)), _solutionFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempProjectPath, String.Format("{0}.csproj", Options.AssemblyName)), _projectFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempProjectPath, "Addin.cs"), _addinFile, Encoding.UTF8);
            if(_applicationFound)
                File.AppendAllText(Path.Combine(TempProjectPath, String.Format("{0}.csproj.user", Options.AssemblyName)), _projectUserFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "AssemblyInfo.cs"), _assemblyFile, Encoding.UTF8);
            if (Options.UseRibbonUI)
                 File.AppendAllText(Path.Combine(TempProjectPath, "RibbonUI.xml"), _ribbonFile, Encoding.UTF8);
            if (Options.UseTaskPane)
            {
                File.AppendAllText(Path.Combine(TempProjectPath, "MyTaskPane.cs"), _taskPaneFile, Encoding.UTF8);
                File.AppendAllText(Path.Combine(TempProjectPath, "MyTaskPane.Designer.cs"), _taskPaneDesignerFile, Encoding.UTF8);
            }
        }

        #endregion
    }
}
