using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.ProjectConverters
{
    internal class ClassLibraryConverterCS : Converter
    {
        #region Fields

        private string _solutionFile;
        private string _projectFile;
        private string _classFile;
        private string _assemblyFile;
        private Guid   _projectGuid;      

        #endregion

        #region Ctor

        internal ClassLibraryConverterCS(ProjectOptions options) : base(options)
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

            _classFile = _classFile.Replace("$safeprojectname$", Options.AssemblyName);
            _classFile = _classFile.Replace("$usingItems$", this.GetNetOfficeProjectUsingItems());

            _assemblyFile = _assemblyFile.Replace("$safeprojectname$", Options.AssemblyName);
            _assemblyFile = _assemblyFile.Replace("$safeprojectdescription$", Options.AssemblyDescription);
            _assemblyFile = _assemblyFile.Replace("$assemblyguid$", Guid.NewGuid().ToString().ToUpper());
        }

        private void ReadRessourceFiles()
        {
            _solutionFile = ReadProjectTemplateFile("ClassLibraryCS.Solution.txt");
            _projectFile = ReadProjectTemplateFile("ClassLibraryCS.Project.txt");
            _classFile = ReadProjectTemplateFile("ClassLibraryCS.Class1.txt");
            _assemblyFile = ReadProjectTemplateFile("ClassLibraryCS.AssemblyInfo.txt");
        }

        private void WriteResultFilesToTempFolder()
        {            
            File.AppendAllText(Path.Combine(TempSolutionPath, String.Format("{0}.sln", Options.AssemblyName)), _solutionFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempProjectPath, String.Format("{0}.csproj", Options.AssemblyName)), _projectFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempProjectPath, "Class1.cs"), _classFile, Encoding.UTF8);
            File.AppendAllText(Path.Combine(TempPropertiesPath, "AssemblyInfo.cs"), _assemblyFile, Encoding.UTF8);
        }

        #endregion
    }
}
