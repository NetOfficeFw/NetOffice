using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OutlookApi.Tools.Contribution
{
    /// <summary>
    /// Outlook dialog related utils
    /// </summary>
    public class OutlookDialogUtils : NetOffice.OfficeApi.Tools.Contribution.DialogUtils
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        internal OutlookDialogUtils(CommonUtils owner) : base(owner)
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// Host application is currently visible
        /// </summary>
        public bool HostVisible
        {
            get
            {
                return this.Owner.TryGetApplicationVisible(false);
            }
        }

        #endregion
    }
}