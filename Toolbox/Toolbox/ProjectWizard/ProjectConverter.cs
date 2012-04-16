using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Text;
using System.Linq;
using System.Xml.Linq;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Zip.Compression.Streams;

namespace NetOffice.DeveloperToolbox
{
    static class ProjectConverter
    {
        static ProjectOptions _projectOptions;
        static List<Control> _listControls;

        public static string ConvertProjectTemplate(ProjectOptions projectOptions, List<Control> listControls)
        {
            _projectOptions = projectOptions;
            _listControls = listControls;

            if (!Directory.Exists(projectOptions.Folder))
                    Directory.CreateDirectory(projectOptions.Folder);
            string templateFolder = Path.Combine(Application.StartupPath, "Project Wizard\\Templates");
            string assemblyFolder = Path.Combine(Application.StartupPath, "Project Wizard\\NetOffice Assemblies");
              
            string fullTemplatePath = Path.Combine(templateFolder, GetTargetTemplate(projectOptions));
            if (!File.Exists(fullTemplatePath))
                throw new System.IO.FileNotFoundException(string.Format("File not found:{0}", fullTemplatePath));

            string targetFolder = Path.Combine(projectOptions.Folder, GetProjectName());
            targetFolder = Path.Combine(targetFolder, GetProjectName());

            if (Directory.Exists(Path.Combine(projectOptions.Folder, GetProjectName())))
                Directory.Delete(Path.Combine(projectOptions.Folder, GetProjectName()), true);
            Directory.CreateDirectory(Path.Combine(projectOptions.Folder, GetProjectName()));
            
            if (Directory.Exists(targetFolder))
                Directory.Delete(targetFolder, true);
            Directory.CreateDirectory(targetFolder);

            FastZip fz = new FastZip();
            fz.ExtractZip(fullTemplatePath, targetFolder, "");            
            DeleteNonUsedFiles(targetFolder);
            RenameProjectFile(targetFolder);
            DoReplace(targetFolder);
            CopyAssemblies(GetAssemblies(), targetFolder, assemblyFolder);
            string projectGuid = ValidateProjectFile(targetFolder);
            CreateSolutionFile(Path.Combine(projectOptions.Folder, GetProjectName()), projectGuid);
            return Path.Combine(projectOptions.Folder, GetProjectName());
        }

        static void CreateSolutionFile(string targetFolder, string projectGuid)
        {
            string solutionContent = CodeTemplates.SolutionFile(_projectOptions.Language, (_projectOptions.IDE == IDE.VS2010));
            solutionContent = solutionContent.Replace("%ProjectName%", GetProjectName());
            solutionContent = solutionContent.Replace("%ProjectGUID%", projectGuid);
            string filePath = Path.Combine(targetFolder, GetProjectName() + ".sln");
            File.AppendAllText(filePath, solutionContent, Encoding.UTF8);
        }

        private static string ValidateProjectFile(string targetFolder)
        {
            string guid = Guid.NewGuid().ToString().ToUpper();
            string extension = _projectOptions.Language == ProgrammingLanguage.CSharp ? ".csproj" : ".vbproj";

            string[] files = Directory.GetFiles(targetFolder, "*" + extension, SearchOption.AllDirectories);
            foreach (string file in files)
            {
                XDocument document = XDocument.Load(file);
                XElement rootNode = (document.FirstNode as XElement);
                rootNode.Attribute("ToolsVersion").Value = _projectOptions.NetRuntime == "4.0" ? "4.0" : "3.5";
                var properties = (from a in rootNode.Elements() select a);

                foreach (var property in properties)
                {
                    foreach (XElement item in property.Elements())
                    {
                        if (item.Name.LocalName == "ProjectGuid")
                            item.Value = "{" + guid + "}";
                    }
                }
               
                document.Save(file);
            }

            return guid;
        }

        public static void RenameProjectFile(string targetFolder)
        {
            string extension = _projectOptions.Language == ProgrammingLanguage.CSharp ? ".csproj" : ".vbproj";

            string[] files = Directory.GetFiles(targetFolder, "*" + extension, SearchOption.AllDirectories);
            foreach (string file in files)
            {
                string path = Path.GetDirectoryName(file);
                string fileName = Path.GetFileName(file);
             
                string newFileName = GetProjectName() + extension;
                File.Move(file, Path.Combine(path, newFileName));
            }
        }

        private static void DeleteNonUsedFiles(string targetFolder)
        {
            string iconFile = Path.Combine(targetFolder, "__TemplateIcon.ico");
            string vsTemplateFile = Path.Combine(targetFolder, "MyTemplate.vstemplate");
            if (File.Exists(iconFile))
                File.Delete(iconFile);
            if (File.Exists(vsTemplateFile))
                File.Delete(vsTemplateFile);
        }

        private static void CopyAssemblies(string[] assemblies, string targetFolder, string assemblyFolder)
        {
            string targetAssemblyFolder = Path.Combine(targetFolder, "NetOffice");
            Directory.CreateDirectory(targetAssemblyFolder);
            foreach (string sourceAssembly in assemblies)
            {
                string fullSourcePathAssembly = Path.Combine(assemblyFolder, _projectOptions.NetRuntime + "\\" + sourceAssembly + ".dll");
                string fullSourcePathDocuFile = Path.Combine(assemblyFolder, "Documentation\\" + sourceAssembly + ".xml");

                string fullTargetPathAssembly = Path.Combine(targetAssemblyFolder, sourceAssembly + ".dll");
                string fullTargetPathDocuFile = Path.Combine(targetAssemblyFolder, sourceAssembly + ".xml");

                File.Copy(fullSourcePathAssembly, fullTargetPathAssembly);
                File.Copy(fullSourcePathDocuFile, fullTargetPathDocuFile);
            }
        }

        private static void DoReplace(string targetFolder)
        {
            string safeRandomGuid = Guid.NewGuid().ToString().ToUpper();

            if (IsAddinProject())
                DeleteNonUsedAddinFiles(targetFolder);

            string[] files = Directory.GetFiles(targetFolder, "*.*", SearchOption.AllDirectories);
            foreach (string file in files)
            {
                string fileContent = File.ReadAllText(file, Encoding.UTF8);
                fileContent = fileContent.Replace("$safeprojectname$", GetProjectName());
                fileContent = fileContent.Replace("$targetframeworkversion$", GetFrameworkVersion());
                fileContent = fileContent.Replace("$assemblyReferences$", GetAssemblyReferences());
                fileContent = fileContent.Replace("$usingItems$", GetUsings());
                fileContent = fileContent.Replace("$randomGuid$", safeRandomGuid);
                fileContent = fileContent.Replace("$safeitemname$", "Addin");
                fileContent = fileContent.Replace("$assemblyGuid$", Guid.NewGuid().ToString());

                if (IsAddinProject())
                {
                    if (IsRibbonUIEnabled())
                    {
                        fileContent = fileContent.Replace("$ribbonFileReference$", CodeTemplates.RibbonReference);
                        fileContent = fileContent.Replace("$ribbonImplement$", CodeTemplates.RibbonImplement(_projectOptions.Language));
                        string ribbonImplement = CodeTemplates.RibbonImplementCode(_projectOptions.Language)+ "\r\n";
                        fileContent = fileContent.Replace("$ribbonUIImplementMethod$", ribbonImplement);
                        fileContent = fileContent.Replace("$helperCode$", CodeTemplates.HelperCode(_projectOptions.Language));
                    }
                    else
                    {
                        fileContent = fileContent.Replace("$ribbonFileReference$", "");
                        fileContent = fileContent.Replace("$ribbonImplement$", "");
                        fileContent = fileContent.Replace("$ribbonUIImplementMethod$", "");
                        fileContent = fileContent.Replace("$helperCode$", "");
                    }

                    if (IsClassicUIEnabled())
                    {
                        fileContent = fileContent.Replace("$classicUICreateCall$", CodeTemplates.ClassicUICall(_projectOptions.Language));
                        fileContent = fileContent.Replace("$classicUIRemoveCall$", CodeTemplates.ClassicUIRemoveCall(_projectOptions.Language));
                        fileContent = fileContent.Replace("$classicUICreateRemoveMethod$",  CodeTemplates.ClassicUIMethod(_projectOptions.Language) + CodeTemplates.ClassicUIRemoveMethod(_projectOptions.Language));
                    }
                    else
                    {
                        fileContent = fileContent.Replace("$classicUICreateCall$", "");
                        fileContent = fileContent.Replace("$classicUIRemoveCall$", "");
                        fileContent = fileContent.Replace("$classicUICreateRemoveMethod$", "");
                    }

                    fileContent = fileContent.Replace("$registerCode$", GetRegisterCode());
                    fileContent = fileContent.Replace("$unregisterCode$", GetUnRegisterCode());

                    if (GetSimpleAssemblies().Length == 1)
                        fileContent = fileContent.Replace("$ApplicationField$", CodeTemplates.AppFieldCode(_projectOptions.Language).Replace("%OfficeApp%", GetSimpleAssemblies()[0]));
                    else
                    {
                        if (_projectOptions.Language == ProgrammingLanguage.CSharp)
                            fileContent = fileContent.Replace("$ApplicationField$", "\t\tCOMObject _application;\r\n");
                        else
                            fileContent = fileContent.Replace("$ApplicationField$", "\tDim _application As COMObject\r\n");
                    }

                    if (GetSimpleAssemblies().Length == 1)
                        fileContent = fileContent.Replace("$ApplicationConstruction$", CodeTemplates.AppConstructionCode(_projectOptions.Language).Replace("%OfficeApp%", GetSimpleAssemblies()[0]));
                    else
                    {
                        if (_projectOptions.Language == ProgrammingLanguage.CSharp)
                            fileContent = fileContent.Replace("$ApplicationConstruction$", "\t\t\t_application = Factory.CreateObjectFromComProxy(null, Application);");
                        else
                            fileContent = fileContent.Replace("$ApplicationConstruction$", "\t\t_application = Factory.CreateObjectFromComProxy(Nothing, Application)");
                    }
                    fileContent = fileContent.Replace("$ApplicationDestroy$", CodeTemplates.AppDestroyCode(_projectOptions.Language));
                   
                    if(_projectOptions.Language == ProgrammingLanguage.CSharp)
                        fileContent = fileContent.Replace("void IDTExtensibility2.OnStartupComplete(ref Array custom)","\tvoid IDTExtensibility2.OnStartupComplete(ref Array custom)");
                }

                File.Delete(file);
                File.AppendAllText(file, fileContent, Encoding.UTF8);
            }
        }

        private static string GetHiveKey()
        {
            foreach (Control item in _listControls)
            {
                LoadControl loadControl = item as LoadControl;
                if (null != loadControl)
                {
                    return loadControl.Hivekey;
                }
            }
            throw new ArgumentOutOfRangeException("LoadControl");
        }

        private static string GetRegisterCode()
        {
            string result = "";
            foreach (Control item in _listControls)
            {
                HostControl hostControl = item as HostControl;
                if (null != hostControl)
                {
                    List<string> hostApps = hostControl.HostApplications;
                    foreach (string app in hostApps)
                    {
                        string registerCode = CodeTemplates.RegisterCode(_projectOptions.Language);
                        registerCode = registerCode.Replace("%OfficeApp%", app).Replace("%HiveKey%", GetHiveKey()).Replace("%OfficAddinKey%", GetOfficeAddinKey(app));
                        registerCode = registerCode.Replace("%Name%", GetProjectName()).Replace("%Description%", GetProjectDescription()).Replace("%LoadBehavior%", GetLoadBehaviour());
                        result += registerCode;
                    }
                    return result;
                }
            }
            throw new ArgumentOutOfRangeException("HostControl");
        }

        private static string GetUnRegisterCode()
        {
            string result = "";
            foreach (Control item in _listControls)
            {
                HostControl hostControl = item as HostControl;
                if (null != hostControl)
                {
                    List<string> hostApps = hostControl.HostApplications;
                    foreach (string app in hostApps)
                    {
                        string unRegisterCode = CodeTemplates.UnRegisterCode(_projectOptions.Language);
                        unRegisterCode = unRegisterCode.Replace("%HiveKey%", GetHiveKey()).Replace("%OfficAddinKey%", GetOfficeAddinKey(app));
                        result += unRegisterCode;
                    }
                    return result;
                }
            }
            throw new ArgumentOutOfRangeException("HostControl");
        }

        private static string GetOfficeAddinKey(string officeApp)
        {
            return "Software\\Microsoft\\Office\\" + officeApp + "\\Addins\\" + GetProjectName() + ".Addin";
        }

        private static void DeleteNonUsedAddinFiles(string targetFolder)
        {
            if (!IsRibbonUIEnabled())
            {
                string targetFile = Path.Combine(targetFolder, "RibbonUI.xml");
                if (File.Exists(targetFile))
                    File.Delete(targetFile);   
            }
        }

        private static bool IsClassicUIEnabled()
        {
            foreach (Control item in _listControls)
            {
                GuiControl guiControl = item as GuiControl;
                if (null != guiControl)
                {
                    return guiControl.ClassicUIEnabled;
                }
            }
            throw new ArgumentOutOfRangeException("GuiControl");
        }

        private static bool IsRibbonUIEnabled()
        {
            foreach (Control item in _listControls)
            {
                GuiControl guiControl = item as GuiControl;
                if (null != guiControl)
                {
                    return guiControl.RibbonUIEnabled;
                }
            }
            throw new ArgumentOutOfRangeException("GuiControl");
        }

        private static bool IsAddinProject()
        {
            return (_projectOptions.ProjectType == ProjectType.Addin);
        }

        private static string GetLoadBehaviour()
        {
            foreach (Control item in _listControls)
            {
                LoadControl loadControl = item as LoadControl;
                if (null != loadControl)
                {
                    return loadControl.LoadBehaviour;
                }
            }
            throw new ArgumentOutOfRangeException("LoadControl");
        }

        private static string GetRibbonFileReferences()
        {
            foreach (Control item in _listControls)
            {
                NameControl nameControl = item as NameControl;
                if (null != nameControl)
                {
                    return nameControl.AssemblyName;
                }
            }
            throw new ArgumentOutOfRangeException("NameControl");
        }

        private static string[] GetSimpleAssemblies()
        {
            foreach (Control item in _listControls)
            {
                HostControl hostControl = item as HostControl;
                if (null != hostControl)
                {
                    List<string> hostApps = hostControl.HostApplications;
                    return hostApps.ToArray();
                }
            }
            throw new ArgumentOutOfRangeException("HostControl");
        }

        private static string[] GetAssemblies()
        {
            foreach (Control item in _listControls)
            {
                HostControl hostControl = item as HostControl;
                if (null != hostControl)
                {
                    List<string> hostApps = hostControl.HostApplications;
                    AddDependenciesToList(hostApps);
                    for (int i = 0; i < hostApps.Count; i++)
                        hostApps[i] = hostApps[i] + "Api";
                    hostApps.Add("LateBindingApi.Core");
                    return hostApps.ToArray();
                }
            }
            throw new ArgumentOutOfRangeException("HostControl");
        }

        private static string GetUsings()
        {
            string result = "";
            foreach (Control item in _listControls)
            {
                HostControl hostControl = item as HostControl;
                if (null != hostControl)
                {
                    List<string> hostApps = hostControl.HostApplications;
                    AddDependenciesToList(hostApps);

                    if (_projectOptions.Language == ProgrammingLanguage.CSharp)
                        result += "using LateBindingApi.Core;" + Environment.NewLine;
                    else
                        result += "Imports LateBindingApi.Core" + Environment.NewLine;

                    foreach (string app in hostApps)
                        result += CodeTemplates.Using(_projectOptions.Language).Replace("%Alias%", app).Replace("%Name%", app) + Environment.NewLine;

                    return result;
                }
            }
            throw new ArgumentOutOfRangeException("HostControl");
        }

        private static string GetAssemblyReferences()
        {
            string result = "";
            foreach (Control item in _listControls)
            {
                HostControl hostControl = item as HostControl;
                if (null != hostControl)
                {
                    List<string> hostApps = hostControl.HostApplications;
                    AddDependenciesToList(hostApps);
                    foreach (string app in hostApps)
                    {
                        result += CodeTemplates.AssemblyReference.Replace("%Name%", app).Replace("%RealName%", app + "Api") + Environment.NewLine;
                    }
                    result += CodeTemplates.AssemblyReference.Replace("%Name%", "LateBindingApi").Replace("%RealName%", "LateBindingApi.Core") + Environment.NewLine;

                    return result;
                }
            }
            throw new ArgumentOutOfRangeException("HostControl");
        }

        private static void AddDependenciesToList(List<string> hostApps)
        {
            List<string> addList = new List<string>();
            foreach (string  name in hostApps)
            {
                 switch (name)
                {
                    case "Excel":
                        AddToList(addList, "Office");
                        AddToList(addList, "VBIDE");
                        break;
                    case "Word":
                        AddToList(addList, "Office");
                        AddToList(addList, "VBIDE");
                        break;
                    case "Outlook":
                        AddToList(addList, "Office");
                        break;
                    case "PowerPoint":
                        AddToList(addList, "Office");
                        AddToList(addList, "VBIDE");
                        break;
                    case "Access":
                        AddToList(addList, "Office");
                        AddToList(addList, "DAO");
                        AddToList(addList, "VBIDE");
                        AddToList(addList, "ADODB");
                        AddToList(addList, "OWC10");
                        AddToList(addList, "MSDATASRC");
                        AddToList(addList, "MSComctlLib");
                        break;
                    default:
                        break;
                }
            }
            foreach (string item in addList)
            {
                hostApps.Add(item);
            }
        }

        private static void AddToList(List<string> list, string name)
        {
            foreach (string item in list)
            {
                if (item == name)
                    return;
            }

            list.Add(name);
        }

        private static string GetFrameworkVersion()
        {
            return _projectOptions.NetRuntime;
        }

        private static string GetProjectDescription()
        {
            foreach (Control item in _listControls)
            {
                NameControl nameControl = item as NameControl;
                if (null != nameControl)
                {
                    return nameControl.AssemblyDescription;
                }
            }
            throw new ArgumentOutOfRangeException("NameControl");
        }

        private static string GetProjectName()
        {
            foreach (Control item in _listControls)
            {
                NameControl nameControl = item as NameControl;
                if (null != nameControl)
                {
                    return nameControl.AssemblyName;
                }
            }
            throw new ArgumentOutOfRangeException("NameControl");
        }

        private static string GetTargetTemplate(ProjectOptions projectOptions)
        {
            switch (projectOptions.Language)
            { 
                case ProgrammingLanguage.CSharp:
                    switch (projectOptions.ProjectType)
                    { 
                        case ProjectType.Addin:
                            return "Automation Addin C#.zip";
                        case ProjectType.WindowsForms:
                            return "WindowsForms Application C#.zip";
                        case ProjectType.ClassLibrary:
                            return "ClassLibrary C#.zip";
                        default: // Console
                            return "Console Application C#.zip";
                    }
                default: // VB
                    switch (projectOptions.ProjectType)
                    { 
                        case ProjectType.Addin:
                            return "Automation Addin VB.zip";
                        case ProjectType.WindowsForms:
                            return "WindowsForms Application VB.zip";
                        case ProjectType.ClassLibrary:
                            return "ClassLibrary VB.zip";
                        default: // Console
                            return "Console Application VB.zip";
                    }
            }
        }
    }
}
