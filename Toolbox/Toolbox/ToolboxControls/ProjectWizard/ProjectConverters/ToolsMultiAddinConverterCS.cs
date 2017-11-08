using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.ProjectConverters
{
    internal class ToolsMultiAddinConverterCS : Converter
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

        internal ToolsMultiAddinConverterCS(ProjectOptions options) : base(options)
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

            _addinFile = _addinFile.Replace("$appName$", "Office");
            _addinFile = _addinFile.Replace("$safeprojectname$", Options.AssemblyName);
            _addinFile = _addinFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingToolsItems());
            _addinFile = _addinFile.Replace("$randomGuid$", Guid.NewGuid().ToString().ToUpper());
            _addinFile = _addinFile.Replace("$name$", Options.AssemblyName);
            _addinFile = _addinFile.Replace("$description$", Options.AssemblyDescription);
            _addinFile = _addinFile.Replace("$loadbehavior$", ConvertLoadBehavoir(Options.LoadBehaviour));

            _assemblyFile = _assemblyFile.Replace("$safeprojectname$", Options.AssemblyName);
            _assemblyFile = _assemblyFile.Replace("$safeprojectdescription$", Options.AssemblyDescription);
            _assemblyFile = _assemblyFile.Replace("$assemblyguid$", Guid.NewGuid().ToString().ToUpper());

            _taskPaneFile = _taskPaneFile.Replace("$safeprojectname$", Options.AssemblyName);
            _taskPaneFile = _taskPaneFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingItems());
            _taskPaneFile = _taskPaneFile.Replace("$appName$", "Office");

            _taskPaneDesignerFile = _taskPaneDesignerFile.Replace("$safeprojectname$", Options.AssemblyName);
            _taskPaneDesignerFile = _taskPaneDesignerFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingItems());

            string attribute2String = "\t[MultiRegister(";
            foreach (var item in Options.OfficeApps)
                attribute2String += "RegisterIn." + item + ", ";
            attribute2String = attribute2String.Substring(0, attribute2String.Length -2);
            attribute2String += ")]";
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
                 attributeString += ", CustomUI(\"RibbonUI.xml\", true)".Replace("$safeprojectname$", Options.AssemblyName);
                _addinFile = _addinFile.Replace("$ribbonProperty$", "\t\tinternal Office.IRibbonUI RibbonUI { get; private set; }");
                _projectFile = _projectFile.Replace("$ribbonFileReference$", "  <ItemGroup>\r\n    <EmbeddedResource Include=\"RibbonUI.xml\" />\r\n  </ItemGroup>");
                _ribbonFile = _ribbonFile.Replace("$safeprojectname$", Options.AssemblyName);
                _addinFile = _addinFile.Replace("$ribbonLoad$", ReadProjectTemplateFile("ToolsMultiAddinCS.RibbonImplement.txt"));
                _addinFile = _addinFile.Replace("$safeprojectname$", Options.AssemblyName);
            }
            else
            {
                _addinFile = _addinFile.Replace("$ribbonDefine$", String.Empty);
                _addinFile = _addinFile.Replace("$ribbonProperty$", String.Empty);
                _addinFile = _addinFile.Replace("$ribbonImplement$", String.Empty);
                _projectFile = _projectFile.Replace("$ribbonFileReference$", String.Empty);
                _addinFile = _addinFile.Replace("$ribbonLoad$", String.Empty);
            }

            if (Options.UseTaskPane)
            {
                if (Options.UseRibbonUI || Options.HiveKey == "LocalMachine")
                    attributeString += ", CustomPane(typeof(MyTaskPane), \"My TaskPane\", true, PaneDockPosition.msoCTPDockPositionRight)";
                else
                    attributeString += "CustomPane(typeof(MyTaskPane), \"My TaskPane\", true, PaneDockPosition.msoCTPDockPositionRight)";
            }
            else
            {
                _addinFile = _addinFile.Replace("$customPane$", String.Empty);
            }

            if (!String.IsNullOrWhiteSpace(attributeString))
                attributeString = "\t[" + attributeString + "]";
            _addinFile = _addinFile.Replace("$attributes$", attributeString);

            string getVersion = "\t\t\tConsole.WriteLine(\"Addin started in {0}\", Application.InstanceFriendlyName);";
            _addinFile = _addinFile.Replace("$getversion$", getVersion);

            if (Options.UseClassicUI)
            {
                _addinFile = _addinFile.Replace("$classicUICreateCall$", "\t\t\tCreateUserInterface();");
                _addinFile = _addinFile.Replace("$classicUIRemoveCall$", "\t\t\tRemoveUserInterface();");

                string template = "" +
                                  "\t\tprivate void CreateUserInterface()\r\n\t\t{\r\n            \r\n\t\t}\r\n\r\n" +
                                  "\t\tprivate void RemoveUserInterface()\r\n\t\t{\r\n            \r\n\t\t}\r\n\r\n" +
                                  "";

                _addinFile = _addinFile.Replace("$classicUICreateRemoveMethod$", template);
            }
            else
            {
                _addinFile = _addinFile.Replace("$classicUICreateCall$\r\n", String.Empty);
                _addinFile = _addinFile.Replace("$classicUIRemoveCall$", String.Empty);
                _addinFile = _addinFile.Replace("$classicUICreateRemoveMethod$", String.Empty);
            }

            if (Options.UseToogle)
            {
                _addinFile = _addinFile.Replace("$tooglecode$", ReadProjectTemplateFile("ToolsMultiAddinCS.RibbonToogleImplement.txt"));
                _taskPaneFile = _taskPaneFile.Replace("$tooglecall$", "\t\t\tif(null != ParentAddin && null != ParentAddin.RibbonUI)\r\n\t\t\t\tParentAddin.RibbonUI.InvalidateControl(\"tooglePaneVisibleButton\");");
                _ribbonFile = _ribbonFile.Replace("$tooglebutton$", "           <toggleButton id=\"tooglePaneVisibleButton\" label=\"TaskPane\" imageMso=\"CreateFormBlankForm\" size=\"large\" getPressed=\"TooglePaneVisibleButton_GetPressed\" onAction=\"TooglePaneVisibleButton_Click\" />");
            }
            else
            {
                _addinFile = _addinFile.Replace("$tooglecode$", String.Empty);
                _taskPaneFile = _taskPaneFile.Replace("$tooglecall$", String.Empty);
                _ribbonFile = _ribbonFile.Replace("\r\n$tooglebutton$", String.Empty);
            }

            _addinFile = ValidateFileContentFormat(_addinFile);
        }

        private void ReadRessourceFiles()
        {
            _taskPaneFile = ReadProjectTemplateFile("ToolsMultiAddinCS.TaskPane.txt");
            _taskPaneDesignerFile = ReadProjectTemplateFile("ToolsMultiAddinCS.TaskPane_Designer.txt");
            _ribbonFile = ReadProjectTemplateFile("ToolsMultiAddinCS.RibbonUI.txt");
            _solutionFile = ReadProjectTemplateFile("ToolsMultiAddinCS.Solution.txt");
            _projectFile = ReadProjectTemplateFile("ToolsMultiAddinCS.Project.txt");
            _projectUserFile = ReadProjectTemplateFile("ToolsMultiAddinCS.Project_User.txt");
            _addinFile = ReadProjectTemplateFile("ToolsMultiAddinCS.Addin.txt");
            _assemblyFile = ReadProjectTemplateFile("ToolsMultiAddinCS.AssemblyInfo.txt");
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
