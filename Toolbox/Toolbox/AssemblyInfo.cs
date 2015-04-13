using System;
using System.Reflection;

namespace NetOffice.DeveloperToolbox
{
    /// <summary>
    /// provides assembly informations
    /// </summary>
    internal static class AssemblyInfo
    {
        private static Assembly _executingAssembly;
        private static object _lock = new object();

        /// <summary>
        /// Title of the Assembly
        /// </summary>
        public static string AssemblyTitle
        {
            get
            {
                object[] attributes = GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != String.Empty)
                        return titleAttribute.Title;
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        /// <summary>
        /// Version of the Assembly
        /// </summary>
        public static string AssemblyVersion
        {
            get
            {
                return GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        /// <summary>
        /// Description of the Assembly
        /// </summary>
        public static string AssemblyDescription
        {
            get
            {
                object[] attributes = GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                    return String.Empty;
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        /// <summary>
        /// Product Markup of the Assembly
        /// </summary>
        public static string AssemblyProduct
        {
            get
            {
                object[] attributes = GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                    return String.Empty;
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        /// <summary>
        /// Copyright Notice of the Assembly
        /// </summary>
        public static string AssemblyCopyright
        {
            get
            {
                object[] attributes = GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                    return String.Empty;
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        /// <summary>
        /// Company Markup of the Assembly
        /// </summary>
        public static string AssemblyCompany
        {
            get
            {
                object[] attributes = GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                    return String.Empty;
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }

        private static Assembly GetExecutingAssembly()
        {
            lock (_lock)
            {
                if (null == _executingAssembly)
                    _executingAssembly = Assembly.GetExecutingAssembly();                
            }
            return _executingAssembly;
        }
    }
}
