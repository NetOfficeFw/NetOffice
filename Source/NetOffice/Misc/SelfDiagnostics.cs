using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Text;
using NetOffice.Tools;

namespace NetOffice.Misc
{
    /// <summary>
    /// Represents a collection with self diagnostic informations
    /// </summary>
    public class SelfDiagnostics : List<SelfDiagnostics.DiagnosticItem>
    {
        /// <summary>
        /// Data item
        /// </summary>
        public class DiagnosticItem
        {
            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="name">name as any</param>
            /// <param name="value">value as any</param>
            public DiagnosticItem(string name, string value)
            {
                Name = name;
                Value = value;
            }

            /// <summary>
            /// Information Name
            /// </summary>
            public string Name { get; private set; }

            /// <summary>
            /// Information Value
            /// </summary>
            public string Value { get; private set; }
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="comAddin">addin base</param>
        public SelfDiagnostics(COMAddinBase comAddin)
        {
            Setup(comAddin);

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public SelfDiagnostics()
        {
            Setup(null);
        }

        private void Setup(COMAddinBase comAddin)
        {
            if (null != comAddin)
            {
                OwnerAssembly = comAddin.GetType().Assembly;

                Add(new DiagnosticItem("---", "Runtime"));
                Add(new DiagnosticItem("LoadingTimeElapsed", comAddin.LoadingTimeElapsed.ToString()));

                if (null != comAddin.AppInstance)
                {
                    Add(new DiagnosticItem("AppInstance", comAddin.AppInstance.InstanceName));
                    if (comAddin.AppInstance.EntityIsAvailable("Version", SupportEntityType.Property))
                    {
                        try
                        {
                            object version = comAddin.AppInstance.Invoker.PropertyGet(comAddin.AppInstance, "Version");
                            Add(new DiagnosticItem("Version", version.ToString()));
                        }
                        catch
                        {
                            ;
                        }                     
                    }
                }

                Add(new DiagnosticItem("---", "Self"));
                Add(new DiagnosticItem("Title", AssemblyTitle));
                Add(new DiagnosticItem("Version", AssemblyVersion));
                Add(new DiagnosticItem("Description", AssemblyDescription));
                Add(new DiagnosticItem("Product", AssemblyProduct));
                Add(new DiagnosticItem("Copyright", AssemblyCopyright));
                Add(new DiagnosticItem("Company", AssemblyCompany));
            }

            Add(new DiagnosticItem("---", "Environment"));
            Add(new DiagnosticItem("Is64BitOperatingSystem", Environment.Is64BitOperatingSystem.ToString()));
            Add(new DiagnosticItem("Is64BitProcess", Environment.Is64BitProcess.ToString()));
            Add(new DiagnosticItem("OSVersion", Environment.OSVersion.ToString()));
            Add(new DiagnosticItem("UserInteractive", Environment.UserInteractive.ToString()));
            Add(new DiagnosticItem("HasShutdownStarted", Environment.HasShutdownStarted.ToString()));

            Add(new DiagnosticItem("---", "AppDomain"));
            Add(new DiagnosticItem("FriendlyName", AppDomain.CurrentDomain.FriendlyName));
            Add(new DiagnosticItem("Id", AppDomain.CurrentDomain.Id.ToString()));
            if(null != AppDomain.CurrentDomain.ApplicationIdentity)
                Add(new DiagnosticItem("ApplicationIdentity", AppDomain.CurrentDomain.ApplicationIdentity.ToString()));

            Add(new DiagnosticItem("---", "Assemblies"));
            foreach (Assembly item in AppDomain.CurrentDomain.GetAssemblies())
            {
                AssemblyName assName = item.GetName();
                string name = assName.Name;
                string version = assName.Version.ToString();
                Add(new DiagnosticItem(name, version));
            }
        }

        private Assembly OwnerAssembly { get; set; }

        private string AssemblyTitle
        {
            get
            {
                object[] attributes = OwnerAssembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != String.Empty)
                        return titleAttribute.Title;
                }
                return System.IO.Path.GetFileNameWithoutExtension(OwnerAssembly.CodeBase);
            }
        }
        
        private string AssemblyVersion
        {
            get
            {
                return OwnerAssembly.GetName().Version.ToString();
            }
        }

        private string AssemblyDescription
        {
            get
            {
                object[] attributes = OwnerAssembly.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                    return String.Empty;
                else
                    return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }
        
        private string AssemblyProduct
        {
            get
            {
                object[] attributes = OwnerAssembly.GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                    return String.Empty;
                else
                    return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        private string AssemblyCopyright
        {
            get
            {
                object[] attributes = OwnerAssembly.GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                    return String.Empty;
                else
                    return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }
        
        private string AssemblyCompany
        {
            get
            {
                object[] attributes = OwnerAssembly.GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                    return String.Empty;
                else
                    return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
    }
}
