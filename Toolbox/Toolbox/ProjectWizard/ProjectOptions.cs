using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Globalization;

namespace NetOffice.DeveloperToolbox
{
    public enum ProjectType
    { 
        Addin = 0,
        WindowsForms =1,
        ClassLibrary = 2,
        Console = 3,
        ToolsAddin = 4
    }

    public enum ProgrammingLanguage
    {
        CSharp = 0,
        VB = 1
    }

    public enum IDE
    {
        VS2008 = 0,
        VS2010 = 1
    }

    public class ProjectOptions
    {
        public ProjectOptions(List<IWizardControl> controls)
        {
            ProjectControl projectControl = GetProjectControl(controls);
            ProjectType = ToProjectType(projectControl.SelectedProjectType(1033), projectControl.UseTools);
            ProjectFolderType = projectControl.SelectedProjectFolderType(1033);
            ProjectFolder = GetSelectedFolder(ProjectFolderType);

            EnvironmentControl envControl = GetEnvironmentControl(controls);
            Language = ToLanguage(envControl.SelectedLanguage);
            IDE = ToIDE(envControl.SelectedIDE);
            NetRuntime = ToRuntime(envControl.SelectedRuntime);
            UseNetRuntimeClient = ToRuntimeUseClient(envControl.SelectedRuntime);

            HostControl hostControl = GetHostControl(controls);
            SetOfficeApps(hostControl);

            NameControl nameControl = GetNameControl(controls);
            AssemblyName = nameControl.AssemblyName;
            AssemblyDescription = nameControl.AssemblyDescription;

            LoadControl loadControl = GetLoadControl(controls);
            LoadBehaviour = Convert.ToInt32(loadControl.LoadBehaviour);
            RegistryKey = loadControl.Hivekey;

            GuiControl guiControl = GetGuiControl(controls);
            UseClassicUI = guiControl.ClassicUIEnabled;
            UseRibbonUI = guiControl.RibbonUIEnabled;
            UseTaskPane = guiControl.TaskPaneEnabled;
        }

        public ProjectType ProjectType { get; private set; }
        public string ProjectFolderType { get; private set; }
        public string ProjectFolder { get; private set; }
        public ProgrammingLanguage Language { get; private set; }
        public IDE IDE { get; private set; }
        public double  NetRuntime { get; private set; }
        public bool UseNetRuntimeClient { get; private set; }
        public string[] OfficeApps { get; private set; }
        public string AssemblyName { get; private set; }
        public string AssemblyDescription { get; private set; }
        public int LoadBehaviour { get; private set; }
        public string RegistryKey { get; private set; }

        public bool UseClassicUI { get; private set; }
        public bool UseRibbonUI { get; private set; }
        public bool UseTaskPane { get; private set; }

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

        private double ToRuntime(string value)
        {
            if (value.IndexOf("Client", StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                return 4.0;
            }
            else
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
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
            string folder10 = "Software\\Microsoft\\VisualStudio\\10.0";
            string folder09 = "Software\\Microsoft\\VisualStudio\\9.0";
            string folderExpress10CS = "Software\\Microsoft\\VCSExpress\\10.0_Config";
            string folderExpress09CS = "Software\\Microsoft\\VCSExpress\\9.0_Config";
            string folderExpress10VB = "Software\\Microsoft\\VBExpress\\10.0_Config";
            string folderExpress09VB = "Software\\Microsoft\\VBExpress\\9.0_Config";

            string folderPath = TryGetRegistryValue(folder10, "VisualStudioProjectsLocation");
            if(null == folderPath)
                folderPath = TryGetRegistryValue(folder09, "VisualStudioProjectsLocation");
            if (null == folderPath)
                folderPath = TryGetRegistryValue(folderExpress10CS, "VisualStudioProjectsLocation");
            if (null == folderPath)
                folderPath = TryGetRegistryValue(folderExpress09CS, "VisualStudioProjectsLocation");
            if (null == folderPath)
                folderPath = TryGetRegistryValue(folderExpress10VB, "VisualStudioProjectsLocation");
            if (null == folderPath)
                folderPath = TryGetRegistryValue(folderExpress09VB, "VisualStudioProjectsLocation");

            if(null == folderPath)
                folderPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

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
            if (value == "2008")
                return IDE.VS2008;
            else
                return IDE.VS2010;
        }

        private ProjectType ToProjectType(string value, bool useTools)
        {
            switch (value)
            {
                case "AutomationAddin":
                {
                    if (useTools)
                        return ProjectType.ToolsAddin;
                    else
                    return ProjectType.Addin;
                }
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
    }
}
