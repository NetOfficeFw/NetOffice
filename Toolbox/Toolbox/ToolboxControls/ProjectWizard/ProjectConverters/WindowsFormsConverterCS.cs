using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.ProjectConverters
{
    internal class WindowsFormsConverterCS : Converter
    {
        #region Fields

        private string _form1File;
        private string _form1DesignerFile;
        private string _programFile;
        private string _ressourceDesignerFile;
        private string _ressourceResFile;
        private string _settingsDesignerFile;
        private string _settingsSettingsFile;
        private string _solutionFile;
        private string _projectFile;
        private string _assemblyFile;
        private Guid   _projectGuid;      

        #endregion

        #region Ctor

        internal WindowsFormsConverterCS(ProjectOptions options) : base(options)
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

            _form1File = _form1File.Replace("$safeprojectname$", Options.AssemblyName);
            _form1File = _form1File.Replace("$usingItems$", this.GetNetOfficeProjectUsingItems());

            _form1DesignerFile = _form1DesignerFile.Replace("$safeprojectname$", Options.AssemblyName);

            _ressourceDesignerFile = _ressourceDesignerFile.Replace("$safeprojectname$", Options.AssemblyName);

            _settingsDesignerFile = _settingsDesignerFile.Replace("$safeprojectname$", Options.AssemblyName);

            _programFile = _programFile.Replace("$safeprojectname$", Options.AssemblyName);

            _solutionFile = _solutionFile.Replace("$safeprojectname$", Options.AssemblyName);
            _solutionFile = _solutionFile.Replace("$projectguid$", _projectGuid.ToString().ToUpper());
            _solutionFile = _solutionFile.Replace("$solutionformat$", this.SolutionFormats[Options.IDE]);
            _solutionFile = _solutionFile.Replace("$ideversion$", this.Environments[Options.IDE, Options.Language]);
            
            _projectFile = _projectFile.Replace("$safeprojectname$", Options.AssemblyName);
            _projectFile = _projectFile.Replace("$projectguid$", _projectGuid.ToString().ToUpper());
            _projectFile = _projectFile.Replace("$toolsversion$", this.Tools[Options.IDE]);
            _projectFile = _projectFile.Replace("$targetframeworkversion$", this.Runtimes[Options.NetRuntime]);
            _projectFile = _projectFile.Replace("$assemblyReferences$", this.GetNetOfficeProjectReferenceItems());

            _assemblyFile = _assemblyFile.Replace("$safeprojectname$", Options.AssemblyName);
            _assemblyFile = _assemblyFile.Replace("$safeprojectdescription$", Options.AssemblyDescription);
            _assemblyFile = _assemblyFile.Replace("$assemblyguid$", Guid.NewGuid().ToString().ToUpper());
        }

        private void ReadRessourceFiles()
        {
            _form1File = ReadProjectTemplateFile("WindowsFormsCS.Form1.txt");
            _form1DesignerFile = ReadProjectTemplateFile("WindowsFormsCS.Form1_Designer.txt");
            _programFile = ReadProjectTemplateFile("WindowsFormsCS.Program.txt");
            _ressourceDesignerFile = ReadProjectTemplateFile("WindowsFormsCS.Resources_Designer.txt");
            _ressourceResFile = ReadProjectTemplateFile("WindowsFormsCS.Resources_resx.txt");
            _settingsDesignerFile = ReadProjectTemplateFile("WindowsFormsCS.Settings_Designer.txt");
            _settingsSettingsFile = ReadProjectTemplateFile("WindowsFormsCS.Settings_settings.txt");
            _solutionFile = ReadProjectTemplateFile("WindowsFormsCS.Solution.txt");
            _projectFile = ReadProjectTemplateFile("WindowsFormsCS.Project.txt");
            _assemblyFile = ReadProjectTemplateFile("WindowsFormsCS.AssemblyInfo.txt");
        }

        private void WriteResultFilesToTempFolder()
        {
            File.AppendAllText(Path.Combine(TempProjectPath, "Form1.cs"), _form1File, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempProjectPath, "Form1.Designer.cs"), _form1DesignerFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempProjectPath, "Program.cs"), _programFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Ressources.Designer.cs"), _ressourceDesignerFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Ressources.resx"), _ressourceResFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Settings.Designer.cs"), _settingsDesignerFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Settings.settings"), _settingsSettingsFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempSolutionPath, String.Format("{0}.sln", Options.AssemblyName)), _solutionFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempProjectPath, String.Format("{0}.csproj", Options.AssemblyName)), _projectFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "AssemblyInfo.cs"), _assemblyFile, Encoding.UTF8);
        }

        #endregion
    }
}
