using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Vbe = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using System.IO.Compression;
using System.Windows.Forms;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace VbeAddin
{
    public class VbaProjectRepository : IDisposable
    {
        [Serializable]
        public class VbaCodeComponents
        {
            public string VbaProjectName { get; set; }

            public VbaCodeComponent[] Components { get; set; }
        }

        [Serializable]
        public class VbaCodeComponent
        {
            public string Name { get; set; }

            public string Code { get; set; }
        }

        protected sealed class DeserializationAppDomainBinder : SerializationBinder
        {
            public override Type BindToType(string assemblyName, string typeName)
            {
                var toAssemblyName = assemblyName.Split(',')[0];
                var result = (from assembly in AppDomain.CurrentDomain.GetAssemblies()
                        where assembly.FullName.Split(',')[0] == toAssemblyName
                        select assembly.GetType(typeName)).FirstOrDefault();
                return result;
            }
        }

        private Vbe.VBE _environment;

        public VbaProjectRepository(Vbe.VBE environment)
        {
            if (null == environment)
                throw new ArgumentNullException("environment");
            _environment = environment.Clone() as Vbe.VBE;
        }

        public IEnumerable<string> ProjectsAvailable()
        {
            string rootPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string folderName = "VbeCodeShare";
            string fullPath = Path.Combine(rootPath, folderName);
            if (Directory.Exists(fullPath))
            {
                List<string> result = new List<string>();
                var files = Directory.GetFiles(fullPath, "*.bin");
                foreach (var item in files)
                    result.Add(Path.GetFileNameWithoutExtension(item));
                return result;                   
            }
            else
                return Enumerable.Empty<string>();
        }

        public void Import(string projectName)
        {
            var target = _environment.ActiveVBProject;
            
            using (Stream file = File.OpenRead(BuildFileName(projectName)))
            {
                var formatter = new BinaryFormatter
                {
                    AssemblyFormat = System.Runtime.Serialization.Formatters.FormatterAssemblyStyle.Simple,
                    Binder = new DeserializationAppDomainBinder()
                };
                VbaCodeComponents result = (VbaCodeComponents)formatter.Deserialize(file);
                foreach (var item in result.Components)
                {
                    Vbe.VBComponent module = target.VBComponents.FirstOrDefault(e => e.Name == item.Name);
                    if (null == module)
                    {
                        module = target.VBComponents.Add(Vbe.Enums.vbext_ComponentType.vbext_ct_StdModule);
                        module.Name = item.Name;
                    }
                    else
                        module.CodeModule.DeleteLines(1, module.CodeModule.CountOfLines);

                    module.CodeModule.AddFromString(item.Code);
                }
            }
        }

        public void Export(string projectName)
        {
            var target = _environment.VBProjects.First(e => e.Name == projectName);
            var formatter = new BinaryFormatter();
            ValidateDirectoryExists();
            using (Stream file = File.Create(BuildFileName(projectName)))
            {
                var components = target.VBComponents;
                VbaCodeComponents result = new VbaCodeComponents();
                result.VbaProjectName = projectName;
                var comps = new List<VbaCodeComponent>();
               
                foreach (var component in components)
                {
                    if (component.Type != Vbe.Enums.vbext_ComponentType.vbext_ct_StdModule)
                        continue;
                    string name = component.Name;
                    string code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines);
                    comps.Add(new VbaCodeComponent() { Name = name, Code = code });
                }
                result.Components = comps.ToArray();
                formatter.Serialize(file, result);
            }
            target.Dispose();
        }

        private void ValidateDirectoryExists()
        {
            string rootPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string folderName = "VbeCodeShare";
            string fullPath = Path.Combine(rootPath, folderName);
            if (!Directory.Exists(fullPath))
                Directory.CreateDirectory(fullPath);
        }

        private string BuildFileName(string projectName)
        {
            string rootPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string folderName = "VbeCodeShare";
            string fileName = projectName + ".bin";
            return Path.Combine(rootPath, folderName, fileName);
        }

        public void Dispose()
        {
            if (null != _environment)
            {
                _environment.Dispose();
                _environment = null;
            }
        }
    }
}