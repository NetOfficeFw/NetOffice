using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Globalization;
using NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard
{
    public enum ProjectType
    { 
        SimpleAddin = 0,
        NetOfficeAddin = 1,
        WindowsForms = 2,
        ClassLibrary = 3,
        Console = 4,        
    }

    public enum ProgrammingLanguage
    {
        CSharp = 0,
        VB = 1
    }

    public enum IDE
    {
        VS2010 = 0,
        VS20131517 = 2
    }
    
    public enum NetVersion
    { 
        Net4 = 0,
        Net4Client = 1,
        Net45 = 2,
        Net451 = 3,
        Net452 = 4,
        Net46 = 5,
        Net461 = 6
    }

    public class ProjectOptions
    {
        #region Ctor

        public ProjectOptions(List<IWizardControl> controls)
        {
            ProjectControl projectControl = GetProjectControl(controls);
            ProjectType = projectControl.SelectedProjectType(); // ToProjectType(projectControl.SelectedProjectType(1033), projectControl.UseTools);
            ProjectFolderType = projectControl.SelectedProjectFolderType();
            ProjectFolder = projectControl.CalculatedFolder;

            EnvironmentControl envControl = GetEnvironmentControl(controls);
            Language = ToLanguage(envControl.SelectedLanguage);
            IDE = ToIDE(envControl.SelectedIDE);
            NetRuntimeTarget = ToRuntime(envControl.SelectedRuntime);
            UseNetRuntimeClient = ToRuntimeUseClient(envControl.SelectedRuntime);

            HostControl hostControl = GetHostControl(controls);
            SetOfficeApps(hostControl);

            NameControl nameControl = GetNameControl(controls);
            AssemblyName = nameControl.AssemblyName;
            AssemblyDescription = nameControl.AssemblyDescription;

            LoadControl loadControl = GetLoadControl(controls);
            LoadBehaviour = Convert.ToInt32(loadControl.LoadBehaviour);

            List<string> list = new List<string>();
            foreach (var item in OfficeApps)
                list.Add(String.Format("Software\\Microsoft\\Office\\{0}\\AddIns", item));
            RegistryKeys = list.ToArray();

            HiveKey = loadControl.Hivekey;

            GuiControl guiControl = GetGuiControl(controls);
            UseClassicUI = guiControl.ClassicUIEnabled;
            UseRibbonUI = guiControl.RibbonUIEnabled;
            UseTaskPane = guiControl.TaskPaneEnabled;
            UseToogle = guiControl.ToogleEnabled;

            if (NetRuntimeTarget == "4.0")
                NetRuntime = UseNetRuntimeClient == true ? NetVersion.Net4Client : NetVersion.Net4;
            else if (NetRuntimeTarget == "4.5")
                NetRuntime = NetVersion.Net45;
            else if (NetRuntimeTarget == "4.5.1")
                NetRuntime = NetVersion.Net451;
            else if (NetRuntimeTarget == "4.5.2")
                NetRuntime = NetVersion.Net452;
            else if (NetRuntimeTarget == "4.6")
                NetRuntime = NetVersion.Net46;
            else if (NetRuntimeTarget == "4.6.1")
                NetRuntime = NetVersion.Net461;
            else
                throw new IndexOutOfRangeException("NetRuntimeTarget");
        }

        #endregion

        #region Properties

        public ProjectType ProjectType { get; private set; }
        public string ProjectFolderType { get; private set; }
        public string ProjectFolder { get; private set; }
        public ProgrammingLanguage Language { get; private set; }
        public IDE IDE { get; private set; }
        public NetVersion NetRuntime { get; private set; }
        public string  NetRuntimeTarget { get; private set; }
        public bool UseNetRuntimeClient { get; private set; }
        public string[] OfficeApps { get; private set; }
        public string AssemblyName { get; private set; }
        public string AssemblyDescription { get; private set; }
        public int LoadBehaviour { get; private set; }
        public string HiveKey { get; private set; }
        public string[] RegistryKeys { get; private set; }

        public bool UseClassicUI { get; private set; }
        public bool UseRibbonUI { get; private set; }
        public bool UseTaskPane { get; private set; }
        public bool UseToogle { get; private set; }

        #endregion

        #region Methods

        private void SetOfficeApps(HostControl control)
        {
            List<string> list = new List<string>();
            if (control.checkBoxExcel.Checked)
                list.Add("Excel");
            if (control.checkBoxWord.Checked)
                list.Add("Word");
            if (control.checkBoxOutlook.Checked)
                list.Add("Outlook");
            if (control.checkBoxPowerPoint.Checked)
                list.Add("PowerPoint");
            if (control.checkBoxAccess.Checked)
                list.Add("Access");
            if (control.checkBoxProject.Checked)
                list.Add("MSProject");
            if (control.checkBoxVisio.Checked)
                list.Add("Visio");

            OfficeApps = list.ToArray();
        }

        private string ToRuntime(string value)
        {
            if (value.IndexOf("Client", StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                return "4.0";
            }
            else
            { 
                return value;
            }
        }

        private bool ToRuntimeUseClient(string value)
        {
            if (value.IndexOf("Client", StringComparison.InvariantCultureIgnoreCase) > -1)
                return true;
            else
                return false;
        }

        private ProgrammingLanguage ToLanguage(string value)
        {
            if (value == "C#")
                return ProgrammingLanguage.CSharp;
            else
                return ProgrammingLanguage.VB;
        }

        private string GetSelectedFolder(string selectedFolder)
        {
            switch (ProjectFolderType)
            {
                case "ApplicationData":
                    return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                case "Desktop":
                    return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                case "User":
                    return Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                case "VSProject":
                    return GetVisualStudioProjectFolder();
                case "Custom":
                    return selectedFolder;
                default:
                    throw new IndexOutOfRangeException("ProjectFolderType");
            }
        }

        internal static string GetVisualStudioProjectFolder()
        {
            string folder11 = "Software\\Microsoft\\VisualStudio\\11.0";
            string folder10 = "Software\\Microsoft\\VisualStudio\\10.0";
            string folder09 = "Software\\Microsoft\\VisualStudio\\9.0";
            string folderExpress11CS = "Software\\Microsoft\\VCSExpress\\10.0_Config";
            string folderExpress10CS = "Software\\Microsoft\\VCSExpress\\10.0_Config";
            string folderExpress09CS = "Software\\Microsoft\\VCSExpress\\9.0_Config";
            string folderExpress11VB = "Software\\Microsoft\\VBExpress\\10.0_Config";
            string folderExpress10VB = "Software\\Microsoft\\VBExpress\\10.0_Config";
            string folderExpress09VB = "Software\\Microsoft\\VBExpress\\9.0_Config";

            string folderPath = TryGetRegistryValue(folder11, "VisualStudioProjectsLocation");
            if(null == folderPath)
                folderPath = TryGetRegistryValue(folder10, "VisualStudioProjectsLocation");
            if (null == folderPath)
                folderPath = TryGetRegistryValue(folder09, "VisualStudioProjectsLocation");
            if (null == folderPath)
                folderPath = TryGetRegistryValue(folderExpress11CS, "VisualStudioProjectsLocation");
            if (null == folderPath)
                folderPath = TryGetRegistryValue(folderExpress11VB, "VisualStudioProjectsLocation");
            if (null == folderPath)
                folderPath = TryGetRegistryValue(folderExpress10CS, "VisualStudioProjectsLocation");
            if (null == folderPath)
                folderPath = TryGetRegistryValue(folderExpress09CS, "VisualStudioProjectsLocation");
            if (null == folderPath)
                folderPath = TryGetRegistryValue(folderExpress10VB, "VisualStudioProjectsLocation");
            if (null == folderPath)
                folderPath = TryGetRegistryValue(folderExpress09VB, "VisualStudioProjectsLocation");

            if(null == folderPath)
                folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);

            return folderPath;
        }

        private static string TryGetRegistryValue(string key, string valueName)
        {
            RegistryKey regKey = Registry.CurrentUser.OpenSubKey(key, false);
            if (null != regKey)
            {
                string regValue = regKey.GetValue(valueName) as string;
                regKey.Close();
                return regValue;
            }
            else
                return null;
        }

        private IDE ToIDE(string value)
        {
            switch (value)
            {
                case "2010":
                default:
                   return IDE.VS20131517;
            }
        }

        private ProjectType ToProjectType(string value, bool useTools)
        {
            switch (value)
            {
                case "NetOffice Addin":
                    return ProjectType.NetOfficeAddin;
                case "Simple Automation Addin":
                case "Einfaches Automation Addin":
                    return ProjectType.SimpleAddin;
                case "WindowsForms":
                    return ProjectType.WindowsForms;
                case "Console":
                    return ProjectType.Console;
                case "ClassLibrary":
                    return ProjectType.ClassLibrary;
                default:
                    throw new IndexOutOfRangeException("value");
            }
        }

        private ProjectControl GetProjectControl(List<IWizardControl> controls)
        {
            foreach (var item in controls)
            {
                ProjectControl ctrl = item as ProjectControl;
                if (null != ctrl)
                    return ctrl;
            }
            throw new IndexOutOfRangeException("controls");
        }

        private EnvironmentControl GetEnvironmentControl(List<IWizardControl> controls)
        {
            foreach (var item in controls)
            {
                EnvironmentControl ctrl = item as EnvironmentControl;
                if (null != ctrl)
                    return ctrl;
            }
            throw new IndexOutOfRangeException("controls");
        }

        private HostControl GetHostControl(List<IWizardControl> controls)
        {
            foreach (var item in controls)
            {
                HostControl ctrl = item as HostControl;
                if (null != ctrl)
                    return ctrl;
            }
            throw new IndexOutOfRangeException("controls");
        }

        private NameControl GetNameControl(List<IWizardControl> controls)
        {
            foreach (var item in controls)
            {
                NameControl ctrl = item as NameControl;
                if (null != ctrl)
                    return ctrl;
            }
            throw new IndexOutOfRangeException("controls");
        }

        private LoadControl GetLoadControl(List<IWizardControl> controls)
        {
            foreach (var item in controls)
            {
                LoadControl ctrl = item as LoadControl;
                if (null != ctrl)
                    return ctrl;
            }
            throw new IndexOutOfRangeException("controls");
        }

        private GuiControl GetGuiControl(List<IWizardControl> controls)
        {
            foreach (var item in controls)
            {
                GuiControl ctrl = item as GuiControl;
                if (null != ctrl)
                    return ctrl;
            }
            throw new IndexOutOfRangeException("controls");
        }

        #endregion
    }
}
