using System;
using System.Collections.Generic;
using System.Drawing;

namespace NetOffice.OfficeApi.Tools.Contribution
{
    /// <summary>
    /// Dialog related utils
    /// </summary>
    public class DialogUtils
    {
        #region Fields

        private const int _currentDefaultLanguage = 1033;
        private CommonUtils _owner;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        protected internal DialogUtils(CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            CurrentLanguage = _currentDefaultLanguage;
            _owner = owner;
            SuppressOnAutomation = true;
            SuppressOnHide = true;
            Layout = new DialogLayoutSettings();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Reference to <see cref="CommonUtils"/> object which owns this instance of DialogUtils class.
        /// </summary>
        public CommonUtils Owner
        {
            get
            {
                return this._owner;
            }
        }

        /// <summary>
        /// Current used language in dialogs. Default is 1033(en-us) If its failed to find a dialog localization set for current language - en-us want be used
        /// </summary>
        public int CurrentLanguage { get; set; }

        /// <summary>
        /// Dont show dialogs if office application is started programmatically for automation , true by default
        /// </summary>
        public bool SuppressOnAutomation { get; set; }

        /// <summary>
        /// Dont show dialogs if office application is currently not visible, true by default
        /// </summary>
        public bool SuppressOnHide { get; set; }

        /// <summary>
        /// Dont show dialogs at all, false by default
        /// </summary>
        public bool SupressGeneraly { get; set; }

        /// <summary>
        /// Default dialogs layout settings
        /// </summary>
        public DialogLayoutSettings Layout { get; private set; }

        #endregion

        #region Methods

        /// <summary>
        /// Returns information show dialogs is currently suspended
        /// </summary>
        /// <returns>true if suspended otherwise false</returns>
        public virtual bool IsCurrentlySuspended()
        {
            if (SupressGeneraly)
                return true;
            if (SuppressOnAutomation && _owner.IsAutomation)
                return true;
            if (SuppressOnHide && false == _owner.TryGetApplicationVisible(true))
                return true;
            return false;
        }

        #endregion
    }
}