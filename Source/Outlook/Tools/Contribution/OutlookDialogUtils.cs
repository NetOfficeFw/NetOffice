﻿using System;
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
                return TryGetApplicationVisible(false);
            }
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Try to detect the visibility of host application main window.
        /// The implementation tries to find a visible Outlook application main window and returns true if
        /// it was found.
        /// </summary>
        /// <param name="defaultResult">fallback result if method fails</param>
        /// <returns>true if application is visible, otherwise false</returns>
        protected override bool TryGetApplicationVisible(bool defaultResult)
        {
            try
            {
                Running.WindowEnumerator enumerator = new Running.WindowEnumerator("rctrl_renwnd32");
                IntPtr[] handles = enumerator.EnumerateWindows(2000);
                if (null != handles)
                {
                    // once again, no linq possible here to keep .Net2 support
                    foreach (IntPtr item in handles)
                    {
                        if (enumerator.IsVisible(item))
                            return true;
                    }
                }
                return false;
            }
            catch (System.Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
                return defaultResult;
            }
        }

        #endregion
    }
}