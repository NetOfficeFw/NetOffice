using System;
using Microsoft.Win32;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ICSharpCode.SharpZipLib;
using ICSharpCode.SharpZipLib.Zip;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.ProjectConverters
{
    /// <summary>
    /// Represents a project converter
    /// </summary>
    internal abstract class Converter : IDisposable
    {
        #region Fields

        private ProjectOptions _options;
        private Guid           _tempGuid;
        private string         _tempPath;

        #endregion

        #region Ctor

        internal Converter(ProjectOptions options)
        {
            _options = options;
            _tempGuid = Guid.NewGuid();
            _tempPath = Path.Combine(Path.GetTempPath(), _tempGuid.ToString());
            Directory.CreateDirectory(_tempPath);
            Directory.CreateDirectory(TempSolutionPath);
            Directory.CreateDirectory(TempProjectPath);
            Directory.CreateDirectory(TempPropertiesPath);
            Directory.CreateDirectory(TempNetOfficePath);
            Environments = new EnvironmenVersions();
            SolutionFormats = new SolutionFormatVersions();
            Tools = new ToolsVersions();
            Runtimes = new RuntimeVersions();
        }

        #endregion

        #region Properties

        protected internal EnvironmenVersions Environments { get; private set; }

        protected internal SolutionFormatVersions SolutionFormats { get; private set; }

        protected internal ToolsVersions Tools { get; private set; }

        protected internal RuntimeVersions Runtimes { get; private set; }

        protected internal ProjectOptions Options
        {
            get
            {
                return _options;
            }
        }

        protected internal string TargetSolutionFile
        {
            get
            {
                string targetFolder = Path.Combine(_options.ProjectFolder, _options.AssemblyName, _options.AssemblyName + ".sln");
                return targetFolder;
            }
        }

        protected internal string TargetSolutionPath
        {
            get
            {
                string targetFolder = Path.Combine(_options.ProjectFolder, _options.AssemblyName);
                return targetFolder;
            }
        }

        protected internal string TargetProjectPath
        {
            get
            {
                string targetFolder = Path.Combine(_options.ProjectFolder, _options.AssemblyName);
                targetFolder = Path.Combine(targetFolder, _options.AssemblyName);
                return targetFolder;
            }
        }

        protected internal string AssembliesFolderPath
        {
            get
            {
                return Program.SubFolder;
            }
        }

        protected internal string TempPath
        {
            get
            {
                return _tempPath;
            }
        }

        protected internal string TempSolutionPath
        {
            get
            {
                string targetFolder = Path.Combine(_tempPath, _options.AssemblyName);
                return targetFolder;
            }
        }

        protected internal string TempProjectPath
        {
            get
            {
                string targetFolder = Path.Combine(_tempPath, _options.AssemblyName, _options.AssemblyName);
                return targetFolder;
            }
        }

        protected internal string TempPropertiesPath
        {
            get
            {
                if (Options.Language == ProgrammingLanguage.CSharp)
                {
                    string targetFolder = Path.Combine(_tempPath, _options.AssemblyName, _options.AssemblyName, "Properties");
                    return targetFolder;
                }
                else
                {
                    string targetFolder = Path.Combine(_tempPath, _options.AssemblyName, _options.AssemblyName, "My Project");
                    return targetFolder;
                }
            }
        }

        protected internal string TempNetOfficePath
        {
            get
            {
                string targetFolder = Path.Combine(_tempPath, _options.AssemblyName, "NetOffice");
                return targetFolder;
            }
        }

        protected internal object TryGetRegistryValue(RegistryHive hive, string path, string valueName = "")
        {
            RegistryKey hiveKey;
            switch (hive)
	        {
		        case RegistryHive.HKEY_Local_Machine:
                    hiveKey = Registry.LocalMachine;
                    break;
                case RegistryHive.HKEY_Current_User:
                    hiveKey = Registry.CurrentUser;
                    break;
                default:
                    throw new ArgumentOutOfRangeException("hive");
	        }

            RegistryKey subKey = hiveKey.OpenSubKey(path);
            if(null != subKey)
            {
                object result =  subKey.GetValue(valueName, null);
                subKey.Close();
                subKey.Dispose();
                return result;
            }
            else
                return null;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Create a Converter instance depending on project options
        /// </summary>
        /// <param name="options">conversion options</param>
        /// <returns></returns>
        public static Converter CreateConverter(ProjectOptions options)
        {
            ValidateOptions(options);
            switch (options.ProjectType)
            {
                case ProjectType.SimpleAddin:
                    {
                        switch (options.Language)
                        {
                            case ProgrammingLanguage.CSharp:
                                if(options.OfficeApps.Length > 1)
                                    return new SimpleMultiAddinConverterCS(options);
                                else
                                    return new SimpleSingleAddinConverterCS(options);
                            case ProgrammingLanguage.VB:
                                if (options.OfficeApps.Length > 1)
                                    return new SimpleMultiAddinConverterVB(options);
                                else
                                    return new SimpleSingleAddinConverterVB(options);
                            default:
                                throw new ArgumentOutOfRangeException("language");
                        }
                    }
                case ProjectType.NetOfficeAddin:
                    {
                        switch (options.Language)
                        {
                            case ProgrammingLanguage.CSharp:
                                if (options.OfficeApps.Length > 1)
                                    return new ToolsMultiAddinConverterCS(options);
                                else
                                    return new ToolsSingleAddinConverterCS(options);
                            case ProgrammingLanguage.VB:
                                if (options.OfficeApps.Length > 1)
                                    return new ToolsMultiAddinConverterVB(options);
                                else
                                    return new ToolsSingleAddinConverterVB(options);
                            default:
                                throw new ArgumentOutOfRangeException("language");
                        }
                    }
                case ProjectType.WindowsForms:
                    {
                        switch (options.Language)
                        {
                            case ProgrammingLanguage.CSharp:
                                return new WindowsFormsConverterCS(options);
                            case ProgrammingLanguage.VB:
                                return new WindowsFormsConverterVB(options);
                            default:
                                throw new ArgumentOutOfRangeException("language");
                        }
                    }
                case ProjectType.ClassLibrary:
                    {
                        switch (options.Language)
                        {
                            case ProgrammingLanguage.CSharp:
                                return new ClassLibraryConverterCS(options);
                            case ProgrammingLanguage.VB:
                                return new ClassLibraryConverterVB(options);
                            default:
                                throw new ArgumentOutOfRangeException("language");
                        }
                    }
                case ProjectType.Console:
                    {
                        switch (options.Language)
                        {
                            case ProgrammingLanguage.CSharp:
                                return new ConsoleConverterCS(options);
                            case ProgrammingLanguage.VB:
                                return new ConsoleConverterVB(options);
                            default:
                                throw new ArgumentOutOfRangeException("language");
                        }
                    }
                default:
                    throw new ArgumentOutOfRangeException("ProjectType");
            }
        }

        /// <summary>
        /// Create new solution
        /// </summary>
        /// <returns>result folder path</returns>
        public abstract string CreateSolution();

        protected internal string ValidateFileContentFormat(string fileContent)
        {
            if (Options.Language == ProgrammingLanguage.CSharp)
            {
                StringBuilder validatedAddinFile = new StringBuilder();
                string[] lines = fileContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                bool lastLineEmpty = false;
                foreach (var item in lines)
                {
                    if (item.Length == 0)
                    {
                        if (false == lastLineEmpty)
                        {
                            validatedAddinFile.AppendLine(item);
                            lastLineEmpty = true;
                        }
                    }
                    else
                    {
                        string tempItem = item.Replace("\t", String.Empty).Replace(" ", String.Empty);
                        if (!String.IsNullOrWhiteSpace(tempItem))
                        {
                            validatedAddinFile.AppendLine(item);
                            lastLineEmpty = false;
                        }
                        else
                            lastLineEmpty = true;

                        if (tempItem == "{")
                            lastLineEmpty = true;
                    }
                }

                fileContent = validatedAddinFile.ToString();
                fileContent = fileContent.Replace("#endregion\r\n\r\n\t}\r\n}", "#endregion\r\n\t}\r\n}");
                return fileContent;
            }
            else
            {
                fileContent = fileContent.Replace("#Region \"", "#Region \"    ");

                StringBuilder validatedAddinFile = new StringBuilder();
                string[] lines = fileContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                bool lastLineEmpty = false;
                foreach (var item in lines)
                {
                    if (item.Length == 0)
                    {
                        if (false == lastLineEmpty)
                        {
                            validatedAddinFile.AppendLine(item);
                            lastLineEmpty = true;
                        }
                    }
                    else
                    {
                        string tempItem = item.Replace("\t", String.Empty).Replace(" ", String.Empty);
                        if (!String.IsNullOrWhiteSpace(tempItem))
                        {
                            validatedAddinFile.AppendLine(item);
                            lastLineEmpty = false;
                        }
                        else
                            lastLineEmpty = true;

                        if (tempItem == "{")
                            lastLineEmpty = true;
                    }
                }



                fileContent = validatedAddinFile.ToString();
                fileContent = fileContent.Replace("#End Region\r\n\r\n\r\n}", "#End Region\r\n\r\n}");
                return fileContent;
            }
        }

        protected internal string ReadProjectTemplateFile(string address)
        {
            return Ressources.RessourceUtils.ReadString("ToolboxControls.ProjectWizard.ProjectTemplates." + address);
        }

        protected internal void MoveTempSolutionFolderToTarget()
        {
            if (!Directory.Exists(TargetSolutionPath))
                Directory.Move(TempSolutionPath, TargetSolutionPath);
            else
                throw new InvalidOperationException(ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1031 ? "Der angegebene Ordner existiert bereits." : "Directory already exists.");
        }

        protected internal string GetNetOfficeProjectReferenceItems()
        {
            string[] officeApps = Options.OfficeApps;

            StringBuilder sb = new StringBuilder();
            string templateItem = "    <Reference Include=\"%Name%, Version=1.7.1.0, Culture=neutral, processorArchitecture=MSIL\">\r\n" +
                                  "      <SpecificVersion>False</SpecificVersion>\r\n" +
                                  "      <HintPath>..\\NetOffice\\%RealName%.dll</HintPath>\r\n" +
                                  "    </Reference>";

            List<string> apps = CreateValidatedReferenceList(officeApps);

            foreach (string app in apps)
                sb.Append(templateItem.Replace("%Name%", app).Replace("%RealName%", app + "Api") + Environment.NewLine);
            sb.Append(templateItem.Replace("%Name%", "NetOffice").Replace("%RealName%", "NetOffice"));

            return sb.ToString();
        }

        protected internal string GetNetOfficeProjectUsingItems()
        {
            string[] officeApps = Options.OfficeApps;
            ProgrammingLanguage language = Options.Language;

            StringBuilder sb = new StringBuilder();

            string usingTemplateCSharp = "using %Alias% = NetOffice.%Name%Api;\r\nusing NetOffice.%Name%Api.Enums;";
            string usingTemplateVB = "Imports %Alias% = NetOffice.%Name%Api\r\nImports NetOffice.%Name%Api.Enums";

            List<string> apps = CreateValidatedReferenceList(officeApps);

            if (language == ProgrammingLanguage.CSharp)
                sb.Append("using NetOffice;" + Environment.NewLine);
            else
                sb.Append("Imports NetOffice" + Environment.NewLine);

            foreach (string app in apps)
            {
                if (language == ProgrammingLanguage.CSharp)
                    sb.Append(usingTemplateCSharp.Replace("%Alias%", app).Replace("%Name%", app) + Environment.NewLine);
                else
                    sb.Append(usingTemplateVB.Replace("%Alias%", app).Replace("%Name%", app) + Environment.NewLine);
            }
            return sb.ToString();
        }

        protected internal string GetNetOfficeProjectUsingToolsItems()
        {
            string[] officeApps = Options.OfficeApps;
            ProgrammingLanguage language = Options.Language;

            StringBuilder sb = new StringBuilder();

            string usingTemplateCSharp = "using %Alias% = NetOffice.%Name%Api;\r\nusing NetOffice.%Name%Api.Enums;";
            string usingTemplateVB = "Imports %Alias% = NetOffice.%Name%Api\r\nImports NetOffice.%Name%Api.Enums";

            List<string> apps = CreateValidatedReferenceList(officeApps);

            if (language == ProgrammingLanguage.CSharp)
            {
                sb.Append("using NetOffice;" + Environment.NewLine);
                sb.Append("using NetOffice.Tools;" + Environment.NewLine);
            }
            else
            { 
                sb.Append("Imports NetOffice" + Environment.NewLine);
                sb.Append("Imports NetOffice.Tools" + Environment.NewLine);
            }

            foreach (string app in apps)
            {
                if (language == ProgrammingLanguage.CSharp)
                {
                    if (IsToolsUsing(app))
                    {
                        sb.Append(usingTemplateCSharp.Replace("%Alias%", app).Replace("%Name%", app) + Environment.NewLine);
                        sb.Append("using NetOffice.%Name%Api.Tools;".Replace("%Name%", app) + Environment.NewLine);
                    }
                    else
                        sb.Append(usingTemplateCSharp.Replace("%Alias%", app).Replace("%Name%", app) + Environment.NewLine);
                }
                else
                {
                    if (IsToolsUsing(app))
                    {
                        sb.Append(usingTemplateVB.Replace("%Alias%", app).Replace("%Name%", app) + Environment.NewLine);
                        sb.Append("Imports NetOffice.%Name%Api.Tools".Replace("%Name%", app) + Environment.NewLine);
                    }
                    else
                        sb.Append(usingTemplateVB.Replace("%Alias%", app).Replace("%Name%", app) + Environment.NewLine);
                }
            }
            return sb.ToString();
        }

        private bool IsToolsUsing(string app)
        {
            switch (app)            
            {
                case "Office":
                case "Excel":
                case "Access":
                case "Outlook":
                case "Word":
                case "MSProject":
                case "PowerPoint":
                    return true;
                default:
                    return false;
            }
        }

        protected internal void CopyUsedNetOfficeAssembliesToTempTarget()
        {
            string[] officeApps = Options.OfficeApps;
            NetVersion runtime = Options.NetRuntime;

            List<string> apps = CreateValidatedReferenceList(officeApps);

            string assembliesFolderPath = AssembliesFolderPath;
            string assembliesTempTarget = TempNetOfficePath;

            File.Copy(Path.Combine(assembliesFolderPath, "NetOffice.xml"), Path.Combine(assembliesTempTarget, "NetOffice.xml"));
            foreach (var item in officeApps)
                File.Copy(Path.Combine(assembliesFolderPath, item + "Api.xml"), Path.Combine(assembliesTempTarget, item + "Api.xml"));

            if (runtime == NetVersion.Net4 || runtime == NetVersion.Net4Client)
            {
                File.Copy(Path.Combine(assembliesFolderPath, "NetOffice.dll"), Path.Combine(assembliesTempTarget, "NetOffice.dll"));
                foreach (var item in apps)
                    File.Copy(Path.Combine(assembliesFolderPath, item + "Api.dll"), Path.Combine(assembliesTempTarget, item + "Api.dll"));
            }
            else
            {
                string targetPackageName = null;
                switch (runtime)
                {
                    case NetVersion.Net2:
                        targetPackageName = Path.Combine(assembliesFolderPath, "2.0.zip");
                        break;
                    case NetVersion.Net3:
                    case NetVersion.Net35:
                        targetPackageName = Path.Combine(assembliesFolderPath, "3.0.zip");
                        break;
                    case NetVersion.Net45:
                        targetPackageName = Path.Combine(assembliesFolderPath, "4.5.zip");
                        break;
                    default:
                        throw new ArgumentOutOfRangeException("runtime");
                }

                using (ZipFile zip = new ZipFile(targetPackageName))
                {
                    Stream streamFirst = zip.GetInputStream(zip.GetEntry("NetOffice.dll"));
                    FileStream fileStreamFirst = File.Create(Path.Combine(assembliesTempTarget, "NetOffice.dll"));
                    streamFirst.CopyTo(fileStreamFirst);
                    fileStreamFirst.Close();

                    foreach (var item in apps)
                    {
                        Stream stream = zip.GetInputStream(zip.GetEntry(item + "Api.dll"));
                        FileStream fileStream = File.Create(Path.Combine(assembliesTempTarget, item + "Api.dll"));
                        stream.CopyTo(fileStream);
                        fileStream.Close();
                    }
                }
            }
        }

        private static List<string> CreateValidatedReferenceList(string[] officeApps)
        {
            List<string> apps = new List<string>();
            List<String> dependecies = new List<string>();
            foreach (var item in officeApps)
                apps.Add(item);

            foreach (var item in apps)
            {
                switch (item)
                {
                    case "Excel":
                        if (!dependecies.Any(a => a == "Office"))
                            dependecies.Add("Office");
                        if (!dependecies.Any(a => a == "VBIDE"))
                            dependecies.Add("VBIDE");
                        break;
                    case "Word":
                        if (!dependecies.Any(a => a == "Office"))
                            dependecies.Add("Office");
                        if (!dependecies.Any(a => a == "VBIDE"))
                            dependecies.Add("VBIDE");
                        break;
                    case "Outlook":
                        if (!dependecies.Any(a => a == "Office"))
                            dependecies.Add("Office");
                        break;
                    case "PowerPoint":
                        if (!dependecies.Any(a => a == "Office"))
                            dependecies.Add("Office");
                        if (!dependecies.Any(a => a == "VBIDE"))
                            dependecies.Add("VBIDE");
                        break;
                    case "Access":
                        if (!dependecies.Any(a => a == "Office"))
                            dependecies.Add("Office");
                        if (!dependecies.Any(a => a == "VBIDE"))
                            dependecies.Add("VBIDE");
                        if (!dependecies.Any(a => a == "DAO"))
                            dependecies.Add("DAO");
                        if (!dependecies.Any(a => a == "ADODB"))
                            dependecies.Add("ADODB");
                        if (!dependecies.Any(a => a == "OWC10"))
                            dependecies.Add("OWC10");
                        if (!dependecies.Any(a => a == "MSDATASRC"))
                            dependecies.Add("MSDATASRC");
                        if (!dependecies.Any(a => a == "MSComctlLib"))
                            dependecies.Add("MSComctlLib");
                        break;
                    case "MSProject":
                        break;
                    case "Visio":
                        if (!dependecies.Any(a => a == "Office"))
                            dependecies.Add("Office");
                        if (!dependecies.Any(a => a == "VBIDE"))
                            dependecies.Add("VBIDE");
                        if (!dependecies.Any(a => a == "MSHTML"))
                            dependecies.Add("MSHTML");
                        break;
                    default:
                        break;
                }
            }

            foreach (var item in dependecies)
                apps.Add(item);

            return apps;
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (Directory.Exists(_tempPath))
                Directory.Delete(_tempPath, true);
        }

        #endregion

        #region Privates

        private static void ValidateOptions(ProjectOptions options)
        {
            switch (options.NetRuntime)
            {
                case NetVersion.Net45:
                    if (options.IDE != IDE.VS2013)
                        throw new ArgumentException("Invalid settings");
                    break;
            }
        }

        #endregion
    }

    internal enum RegistryHive
    { 
        HKEY_Local_Machine = 0,
        HKEY_Current_User
    }

    internal class EnvironmenVersions
    {
        public string this[IDE environment, ProgrammingLanguage language]
        {
            get
            {
                switch (environment)
                {
                    case IDE.VS2010:
                        if (language == ProgrammingLanguage.CSharp)
                            return "# Visual C# Express 2010";
                        else
                            return "# Visual Basic Express 2010";
                    case IDE.VS2012:
                        if (language == ProgrammingLanguage.CSharp)
                            return "# Visual C# Express 2012";
                        else
                            return "# Visual Basic Express 2012";
                    case IDE.VS2013:
                        return "# Visual Studio Express 2013 for Windows Desktop\r\nVisualStudioVersion = 12.0.30723.0\r\nMinimumVisualStudioVersion = 10.0.40219.1";
                    default:
                        throw new ArgumentOutOfRangeException("environment");
                }
            }
        }
    }

    internal class SolutionFormatVersions
    {
        public string this[IDE environment]
        {
            get 
            {
                switch (environment)
                {
                    case IDE.VS2010:
                        return "11.00";
                    case IDE.VS2012:
                        return "11.00";
                    case IDE.VS2013:
                        return "12.00";
                    default:
                        throw new ArgumentOutOfRangeException("environment");
                }
            }
        }
    }

    internal class ToolsVersions
    {
        public string this[IDE environment]
        {
            get
            {
                switch (environment)
                {
                    case IDE.VS2010:
                        return "4.0";
                    case IDE.VS2012:
                        return "4.0";
                    case IDE.VS2013:
                        return "12.00";
                    default:
                        throw new ArgumentOutOfRangeException("environment");
                }
            }
        }
    }

    internal class RuntimeVersions
    {
        public string this[NetVersion runtime]
        {
            get
            {
                switch (runtime)
                {
                    case NetVersion.Net2:
                        return "    <TargetFrameworkVersion>v2.0</TargetFrameworkVersion>";
                    case NetVersion.Net3:
                        return "    <TargetFrameworkVersion>v3.0</TargetFrameworkVersion>";
                    case NetVersion.Net35:
                        return "    <TargetFrameworkVersion>v3.5</TargetFrameworkVersion>";
                    case NetVersion.Net4:
                        return "    <TargetFrameworkVersion>v4.0</TargetFrameworkVersion>";
                    case NetVersion.Net4Client:
                        return "    <TargetFrameworkVersion>v4.0</TargetFrameworkVersion>\r\n    <TargetFrameworkProfile>Client</TargetFrameworkProfile>";
                    case NetVersion.Net45:
                        return "    <TargetFrameworkVersion>v4.5</TargetFrameworkVersion>";
                    default:
                        throw new ArgumentOutOfRangeException("runtime");
                }
            }
        }
    }
}
 