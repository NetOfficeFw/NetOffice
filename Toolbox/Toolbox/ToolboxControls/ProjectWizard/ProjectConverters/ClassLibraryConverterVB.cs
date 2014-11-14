using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.ProjectConverters
{
    internal class ClassLibraryConverterVB : Converter
    {
        #region Fields

        private string _appDesignerFile;
        private string _myApplicationFile;
        private string _ressourceDesgnerFile;
        private string _ressourceResFile;
        private string _settingDesignerFile;
        private string _settingsSettingsFile;

        private string _solutionFile;
        private string _projectFile;
        private string _classFile;
        private string _assemblyFile;
        private Guid   _projectGuid;      

        #endregion

        #region Ctor

        internal ClassLibraryConverterVB(ProjectOptions options) : base(options)
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

            _classFile = _classFile.Replace("$safeprojectname$", Options.AssemblyName);
            _classFile = _classFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingItems());

            _assemblyFile = _assemblyFile.Replace("$safeprojectname$", Options.AssemblyName);
            _assemblyFile = _assemblyFile.Replace("$safeprojectdescription$", Options.AssemblyDescription);
            _assemblyFile = _assemblyFile.Replace("$assemblyguid$", Guid.NewGuid().ToString().ToUpper());
        }

        private void ReadRessourceFiles()
        {
            _appDesignerFile = ReadProjectTemplateFile("ClassLibraryVB.Application_Designer.txt");
            _myApplicationFile = ReadProjectTemplateFile("ClassLibraryVB.Application_myapp.txt");
            _ressourceDesgnerFile = ReadProjectTemplateFile("ClassLibraryVB.Resources_Designer.txt");
            _ressourceResFile = ReadProjectTemplateFile("ClassLibraryVB.Resources_resx.txt");
            _settingDesignerFile = ReadProjectTemplateFile("ClassLibraryVB.Settings_Designer.txt");
            _settingsSettingsFile = ReadProjectTemplateFile("ClassLibraryVB.Settings_settings.txt");
            _solutionFile = ReadProjectTemplateFile("ClassLibraryVB.Solution.txt");
            _projectFile = ReadProjectTemplateFile("ClassLibraryVB.Project.txt");
            _classFile = ReadProjectTemplateFile("ClassLibraryVB.Class1.txt");
            _assemblyFile = ReadProjectTemplateFile("ClassLibraryVB.AssemblyInfo.txt");
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
            File.AppendAllText(Path.Combine(TempProjectPath, "Class1.vb"), _classFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "AssemblyInfo.vb"), _assemblyFile, Encoding.UTF8);
        }

        #endregion
    }
}
