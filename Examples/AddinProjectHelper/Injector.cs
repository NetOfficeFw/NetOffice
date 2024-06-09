using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;

namespace AddinProjectHelper
{
    internal class Injector
    {
        private string[] _targets = new string[] { "01 Simple", "02 Ribbons And Panes", "03 Troubleshooting and Diagnostics", "04 Register and Unregister" };

        internal Injector()
        {
            ExeMapper = new Dictionary<string, string>();
            ExeMapper.Add("Excel", "EXCEL.EXE");
            ExeMapper.Add("Word", "WINWORD.EXE");
            ExeMapper.Add("Outlook", "OUTLOOK.EXE");
            ExeMapper.Add("PowerPoint", "POWERPNT.EXE");
            ExeMapper.Add("Access", "MSACCESS.EXE");
        }

        private Dictionary<string, string> ExeMapper { get; set; }

        internal void Inject(string examplePath, string officeProductName, string officeInstallPath)
        {
            string template = Template();
            string exe = ExeMapper[officeProductName];
            string path = Path.Combine(examplePath, officeProductName);
            string pathCS = Path.Combine(path, @"C#\NetOffice COMAddin Examples");
            string pathVB = Path.Combine(path, @"VB\NetOffice COMAddin Examples");
         
            //c#
            foreach (var target in _targets)
            {
                string projectFile = Path.Combine(pathCS, target, target + ".csproj");
                if (File.Exists(projectFile))
                {
                    string debugPath = Path.Combine(officeInstallPath, exe).Replace("(", "%28").Replace(")", "%29");
                    string content = template.Replace("$Path", debugPath);
                    string userFile = Path.Combine(pathCS, target, target + ".csproj.user");
                    File.WriteAllText(userFile, content, Encoding.UTF8);
                }
            }

            //vb
            foreach (var target in _targets)
            {
                string projectFile = Path.Combine(pathVB, target, target + ".vbproj");
                if (File.Exists(projectFile))
                {
                    string debugPath = Path.Combine(officeInstallPath, exe).Replace("(", "%28").Replace(")", "%29");
                    string content = template.Replace("$Path", debugPath);
                    string userFile = Path.Combine(pathVB, target, target + ".vbproj.user");
                    File.WriteAllText(userFile, content, Encoding.UTF8);
                }
            }
        }

        private static string Template()
        {
            using (var stream = typeof(Injector).Assembly.GetManifestResourceStream("AddinProjectHelper.Template.txt"))
            {
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            } 
        }
    }
}
