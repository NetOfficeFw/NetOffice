using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Informations
{
    /// <summary>
    /// Office Application and NetOffice related diagnostic informations
    /// </summary>
    public class HostInfo : IEnumerable<KeyValuePair<string, string>>
    {      
        #region Fields

        private Utils.CommonUtils _owner;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        internal HostInfo(Utils.CommonUtils owner)
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
                return "Host";
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Do analyze
        /// </summary>
        /// <returns>host info as name=>value collection</returns>
        protected internal virtual IEnumerable<KeyValuePair<string, string>> GetSummary()
        {
            double? appVersion = _owner.TryGetApplicationVersion();
            if (null == appVersion)
                appVersion = 0.0;

            List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>();
            list.Add(new KeyValuePair<string, string>(HeaderCaption, HeaderCaptionLine));
            list.Add(new KeyValuePair<string, string>("Product Name", _owner.OwnerApplication.InstanceName));
            list.Add(new KeyValuePair<string, string>("Product Version", appVersion.ToString()));
            list.Add(new KeyValuePair<string, string>("Proxy Count", _owner.OwnerApplication.Factory.ProxyCount.ToString()));
            list.Add(new KeyValuePair<string, string>("Is Initialized", _owner.OwnerApplication.Factory.IsInitialized.ToString()));
            list.Add(new KeyValuePair<string, string>("Initialized Time MS", _owner.OwnerApplication.Factory.InitializedTime.TotalMilliseconds.ToString()));
            list.Add(new KeyValuePair<string, string>("Loaded Time MS", _owner.Owner.LoadingTimeElapsed.TotalMilliseconds.ToString()));
            list.Add(new KeyValuePair<string, string>("Load Assemblies Unsafe", _owner.OwnerApplication.Factory.Settings.LoadAssembliesUnsafe.ToString()));
            list.Add(new KeyValuePair<string, string>("Operators Enabled", _owner.OwnerApplication.Factory.Settings.EnableOperatorOverlads.ToString()));
            list.Add(new KeyValuePair<string, string>("Management Enabled", _owner.OwnerApplication.Factory.Settings.EnableProxyManagement.ToString()));
            list.Add(new KeyValuePair<string, string>("Safe Enabled", _owner.OwnerApplication.Factory.Settings.EnableSafeMode.ToString()));                       
            list.Add(new KeyValuePair<string, string>("Filter Enabled", _owner.OwnerApplication.Factory.Settings.MessageFilter.Enabled.ToString()));
            list.Add(new KeyValuePair<string, string>("Events Enabled", _owner.OwnerApplication.Factory.Settings.EnableEvents.ToString()));
            list.Add(new KeyValuePair<string, string>("Thread Culture", _owner.OwnerApplication.Factory.Settings.ThreadCulture.LCID.ToString()));
            
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

        /// <summary>
        /// Returns a summary enumerator collect instance properties
        /// </summary>
        /// <returns>summary enumerator</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetSummary().GetEnumerator();
        }

        #endregion
    }
}
