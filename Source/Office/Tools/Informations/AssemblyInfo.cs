using System;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Informations
{
    /// <summary>
    /// Assembly related helper tools (Supports also the internal NetOffice Tools infrastructure)
    /// </summary>
    public class AssemblyInfo : IEnumerable<KeyValuePair<string, string>>
    {
        #region Fields

        private Utils.CommonUtils _owner;
        private Assembly _assembly;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="ownerAssembly">owner Assembly</param>
        public AssemblyInfo(Assembly ownerAssembly)
        {
            if (null == ownerAssembly)
                throw new ArgumentNullException("ownerAssembly");
            _assembly = ownerAssembly;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        internal AssemblyInfo(Utils.CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Title of the assembly
        /// </summary>
        public virtual string AssemblyTitle
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

        /// <summary>
        /// Version of the assembly
        /// </summary>
        public virtual string AssemblyVersion
        {
            get
            {
                return OwnerAssembly.GetName().Version.ToString();
            }
        }

        /// <summary>
        /// Description of the assembly
        /// </summary>
        public virtual string AssemblyDescription
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

        /// <summary>
        /// Product markup of the assembly
        /// </summary>
        public virtual string AssemblyProduct
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

        /// <summary>
        /// Copyright info for the assembly
        /// </summary>
        public virtual string AssemblyCopyright
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

        /// <summary>
        /// Manufactor info for the assembly
        /// </summary>
        public virtual string AssemblyCompany
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

        /// <summary>
        /// Header caption seperator line in summary
        /// </summary>
        private string HeaderCaptionLine
        {
            get
            {
                if (null != _owner)
                    return _owner.HeaderCaptionLine;
                else
                    return Utils.CommonUtils.HeaderCaptionLineDefault;
            }
        }

        /// <summary>
        /// Header caption in summary
        /// </summary>
        protected internal virtual string HeaderCaption
        {
            get
            {
                return "Application";
            }
        }

        /// <summary>
        /// Owner Assembly
        /// </summary>
        private Assembly OwnerAssembly
        {
            get
            {
                if (null != _owner)
                    return _owner.OwnerAssembly;
                if (null == _assembly)
                    _assembly = Assembly.GetExecutingAssembly();
                return _assembly;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Returns a summary enumerator collect instance properties
        /// </summary>
        /// <returns>summary enumerator</returns>
        protected internal virtual IEnumerable<KeyValuePair<string, string>> GetSummary()
        {
            List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>();

            list.Add(new KeyValuePair<string, string>(HeaderCaption, HeaderCaptionLine));
            list.Add(new KeyValuePair<string, string>("Title", AssemblyTitle));
            list.Add(new KeyValuePair<string, string>("Version", AssemblyVersion));
            list.Add(new KeyValuePair<string, string>("Description", AssemblyDescription));
            list.Add(new KeyValuePair<string, string>("Product", AssemblyProduct));
            list.Add(new KeyValuePair<string, string>("Copyright", AssemblyCopyright));
            list.Add(new KeyValuePair<string, string>("Company", AssemblyCompany));

            return list;
        }

        #endregion

        #region IEnumerable<KeyValuePair<string, string>>

        /// <summary>
        /// Returns an enumerator to retrieve the collection
        /// </summary>
        /// <returns>IEnumerator instance</returns>
        public virtual IEnumerator<KeyValuePair<string, string>> GetEnumerator()
        {
            return GetSummary().GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetSummary().GetEnumerator();
        }

        #endregion
    }
}
