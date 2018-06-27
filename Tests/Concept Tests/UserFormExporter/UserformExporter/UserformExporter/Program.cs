using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NetOffice;
using Word = NetOffice.WordApi;

namespace UserformExporter
{
    class Program
    {
        internal static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        internal static void Run()
        {
            NetOffice.Settings.Default.EnableAutomaticQuit = true;

            string fileName = @"C:\Sebastian\NetOffice\Tests\Concept Tests\DynamicsCSharp\Document.docm";
            using (Word.Application application = new NetOffice.WordApi.ApplicationClass())
            {
                using (var doc = application.Documents.Open(fileName))
                {
                    if (doc.HasVBProject)
                    {
                        var project = doc.VBProject;
                        foreach (var component in project.VBComponents)
                        {
                            if (component.Type == NetOffice.VBIDEApi.Enums.vbext_ComponentType.vbext_ct_MSForm)
                            {
                                Console.WriteLine(" - " + component.Name);
                                var exportPath = Path.Combine(Path.GetDirectoryName(fileName), component.Name + ".frm");
                                component.Export(exportPath);
                            }
                        }
                    }
                }
            }
        }
    }
}
