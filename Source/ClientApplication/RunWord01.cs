using System;
using System.IO;
using Word = NetOffice.WordApi;
using VBE = NetOffice.VBIDEApi;

namespace ClientApplication
{
    internal class RunWord01
    {
        internal void Run()
        {
            NetOffice.Settings.Default.EnableAutomaticQuit = true;

            string fileName = @"C:\Sebastian\NetOffice11\RunWord01.docm";
            using (Word.Application application = new NetOffice.WordApi.Application())
            {
                using (var doc = application.Documents.Open(fileName))
                {
                    if (doc.HasVBProject)
                    {
                        var project = doc.VBProject;
                        foreach (VBE.VBComponent component in project.VBComponents)
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
