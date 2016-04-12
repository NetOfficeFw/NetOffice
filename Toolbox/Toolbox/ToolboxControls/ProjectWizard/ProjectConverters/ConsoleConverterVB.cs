using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.ProjectConverters
{
    internal class ConsoleConverterVB : Converter
    {
        #region Fields

        private string _appDesignerFile;
        private string _myApplicationFile;
        private string _ressourceDesignerFile;
        private string _ressourceResFile;
        private string _settingsDesignerFile;
        private string _settingsSettingsFile;
        private string _solutionFile;
        private string _projectFile;
        private string _programFile;
        private string _assemblyFile;
        private Guid   _projectGuid;      

        #endregion

        #region Ctor

        internal ConsoleConverterVB(ProjectOptions options) : base(options)
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

            _ressourceDesignerFile = _ressourceDesignerFile.Replace("$safeprojectname$", Options.AssemblyName);

            _settingsDesignerFile = _settingsDesignerFile.Replace("$safeprojectname$", Options.AssemblyName);

            _solutionFile = _solutionFile.Replace("$safeprojectname$", Options.AssemblyName);
            _solutionFile = _solutionFile.Replace("$projectguid$", _projectGuid.ToString().ToUpper());
            _solutionFile = _solutionFile.Replace("$solutionformat$", this.SolutionFormats[Options.IDE]);
            _solutionFile = _solutionFile.Replace("$ideversion$", this.Environments[Options.IDE, Options.Language]);
            
            _projectFile = _projectFile.Replace("$safeprojectname$", Options.AssemblyName);
            _projectFile = _projectFile.Replace("$projectguid$", _projectGuid.ToString().ToUpper());
            _projectFile = _projectFile.Replace("$toolsversion$", this.Tools[Options.IDE]);
            _projectFile = _projectFile.Replace("$targetframeworkversion$", this.Runtimes[Options.NetRuntime]);
            _projectFile = _projectFile.Replace("$assemblyReferences$", this.GetNetOfficeProjectReferenceItems());

            _programFile = _programFile.Replace("$safeprojectname$", Options.AssemblyName);
            _programFile = _programFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingItems());

            _assemblyFile = _assemblyFile.Replace("$safeprojectname$", Options.AssemblyName);
            _assemblyFile = _assemblyFile.Replace("$safeprojectdescription$", Options.AssemblyDescription);
            _assemblyFile = _assemblyFile.Replace("$assemblyguid$", Guid.NewGuid().ToString().ToUpper());
        }

        private void ReadRessourceFiles()
        {
            _appDesignerFile = ReadProjectTemplateFile("ConsoleVB.Application_Designer.txt");
            _myApplicationFile = ReadProjectTemplateFile("ConsoleVB.Application_myapp.txt");
            _ressourceDesignerFile = ReadProjectTemplateFile("ConsoleVB.Resources_Designer.txt");
            _ressourceResFile = ReadProjectTemplateFile("ConsoleVB.Resources_resx.txt");
            _settingsDesignerFile = ReadProjectTemplateFile("ConsoleVB.Settings_Designer.txt");
            _settingsSettingsFile = ReadProjectTemplateFile("ConsoleVB.Settings_settings.txt");
            _solutionFile = ReadProjectTemplateFile("ConsoleVB.Solution.txt");
            _projectFile = ReadProjectTemplateFile("ConsoleVB.Project.txt");
            _programFile = ReadProjectTemplateFile("ConsoleVB.Program.txt");
            _assemblyFile = ReadProjectTemplateFile("ConsoleVB.AssemblyInfo.txt");
        }

        private void WriteResultFilesToTempFolder()
        {
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Application.Designer.vb"), _appDesignerFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Application.myapp"), _myApplicationFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Resources.Designer.vb"), _ressourceDesignerFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Resources.resx"), _ressourceResFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Settings.Designer.vb"), _settingsDesignerFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "Settings.settings"), _settingsSettingsFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempSolutionPath, String.Format("{0}.sln", Options.AssemblyName)), _solutionFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempProjectPath, String.Format("{0}.vbproj", Options.AssemblyName)), _projectFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempProjectPath,  "Module1.vb"), _programFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "AssemblyInfo.vb"), _assemblyFile, Encoding.UTF8);
        }

        #endregion
    }
}
