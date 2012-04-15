using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox
{
    public enum ProjectType
    { 
        Addin = 0,
        WindowsForms =1,
        ClassLibrary = 2,
        Console = 3
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
        public ProjectOptions(string folder, double netRuntime, ProjectType projectType, ProgrammingLanguage language, IDE ide)
        {
            Folder = folder;
            NetRuntime = netRuntime;
            ProjectType = projectType;
            Language = language;
            IDE = ide;
        }

        public ProjectType ProjectType { get; private set; }
        public ProgrammingLanguage Language { get; private set; }
        public IDE IDE { get; private set; }
        public  double NetRuntime { get; private set; }
        public  string Folder { get; private set; }
    }
}
