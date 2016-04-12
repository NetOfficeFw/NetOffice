using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Informations
{
    /// <summary>
    /// Environment related utils (Supports also the internal NetOffice infrastructure)
    /// </summary>
    public class EnvironmentInfo : IEnumerable<KeyValuePair<string, string>>
    {
        #region Fields

        private Utils.CommonUtils _owner;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        internal EnvironmentInfo(Utils.CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Current Operating System is 64Bit
        /// </summary>
        public bool Is64BitOperatingSystem
        {
            get
            {
                return Environment.Is64BitOperatingSystem;
            }
        }

        /// <summary>
        /// Current Process is 64 Bit
        /// </summary>
        public bool Is64BitProcess
        {
            get
            {
                return Environment.Is64BitProcess;
            }
        }

        /// <summary>
        /// Current Operating System Version
        /// </summary>
        public string OSVersion
        {
            get
            {
                return Environment.OSVersion.ToString();
            }
        }

        /// <summary>
        /// Current Process is running interactive
        /// </summary>
        public bool UserInteractive
        {
            get
            {
                return Environment.UserInteractive;
            }
        }

        /// <summary>
        /// CLR is currently shuting down
        /// </summary>
        public bool HasShutdownStarted
        {
            get
            {
                return Environment.HasShutdownStarted;
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
                return "Environment";
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Do analyze
        /// </summary>
        /// <returns>environment info as name=>value collection</returns>
        protected internal virtual IEnumerable<KeyValuePair<string, string>> GetSummary()
        {
            List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>();

            list.Add(new KeyValuePair<string, string>(HeaderCaption, HeaderCaptionLine));
            list.Add(new KeyValuePair<string, string>("Is64BitOperatingSystem", Is64BitOperatingSystem.ToString()));
            list.Add(new KeyValuePair<string, string>("Is64BitProcess", Is64BitProcess.ToString()));
            list.Add(new KeyValuePair<string, string>("OSVersion", OSVersion.ToString()));
            list.Add(new KeyValuePair<string, string>("UserInteractive", UserInteractive.ToString()));
            list.Add(new KeyValuePair<string, string>("HasShutdownStarted", Environment.HasShutdownStarted.ToString()));

            return list;
        }

        #endregion

        #region IEnumerable<KeyValuePair<string, string>>

        /// <summary>
        /// Returns an enumerator to retrieve the collection
        /// </summary>
        /// <returns>IEnumerator instance</returns>
        IEnumerator<KeyValuePair<string, string>> IEnumerable<KeyValuePair<string, string>>.GetEnumerator()
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
