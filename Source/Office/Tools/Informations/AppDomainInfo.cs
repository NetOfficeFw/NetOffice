using System;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Informations
{
    /// <summary>
    /// AppDomain related helper tools (Supports also the internal NetOffice Tools infrastructure)
    /// </summary>
    public class AppDomainInfo : IEnumerable<KeyValuePair<string, string>>
    {
        #region Fields

        private Utils.CommonUtils _owner;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public AppDomainInfo()
        { 
        
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        internal AppDomainInfo(Utils.CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
        }

        #endregion

        #region Properties

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
                return "Loaded Assemblies";
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Do analyze in current appdomain
        /// </summary>
        /// <returns>loaded assemblies as name=>version collection</returns>
        protected internal virtual IEnumerable<KeyValuePair<string, string>> GetSummary()
        {
            List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>();
            list.Add(new KeyValuePair<string, string>(HeaderCaption, HeaderCaptionLine));
 
            foreach (Assembly item in AppDomain.CurrentDomain.GetAssemblies())
            {
                AssemblyName assName = item.GetName();
                string name = assName.Name;
                string version = assName.Version.ToString();
                list.Add(new KeyValuePair<string, string>(name, version));
            }

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
