﻿using System;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Security.Principal;

namespace NetOffice.OutlookApi.Tools.Contribution
{
    /// <summary>
    /// Various helper for common tasks
    /// </summary>
    public class CommonUtils : NetOffice.OfficeApi.Tools.Contribution.CommonUtils
    {
        #region Fields

        private OutlookApi.Application _ownerApplication;
        private ApplicationUtils _application;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the application
        /// </summary>
        /// <param name="application">owner application</param>
        public CommonUtils(OutlookApi.Application application) : base(application)
        {
            _ownerApplication = application;
        }

        /// <summary>
        /// Creates an instance of the application
        /// </summary>
        /// <param name="application">owner application</param>
        /// <param name="ownerAssembly">owner assembly</param>
        public CommonUtils(OutlookApi.Application application, Assembly ownerAssembly) : base(application, ownerAssembly)
        {
            if (null == application)
                throw new ArgumentNullException("application");

            _ownerApplication = application;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">addin owner instance</param>
        /// <param name="isAutomation">host application is started for automation</param>
        internal CommonUtils(NetOffice.Tools.COMAddinBase owner, bool isAutomation) : base(owner, isAutomation)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">addin owner</param>
        /// <param name="isAutomation">indicates the host application is currently in automation</param>
        /// <param name="ownerAssembly">owner application</param>
        internal CommonUtils(NetOffice.Tools.COMAddinBase owner, bool isAutomation, Assembly ownerAssembly) : base(owner, isAutomation, ownerAssembly)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">addin owner</param>
        /// <param name="ownerType">type information from addin owner</param>
        /// <param name="isAutomation">indicates the host application is currently in automation</param>
        /// <param name="ownerAssembly">owner application</param>
        internal CommonUtils(NetOffice.Tools.COMAddinBase owner, Type ownerType, bool isAutomation, Assembly ownerAssembly) : base(owner, ownerType, isAutomation, ownerAssembly)
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// Application related utils
        /// </summary>
        public ApplicationUtils Application
        {
            get
            {
                if (null == _application)
                    _application = OnCreateApplicationUtils();
                return _application;
            }
        }

        /// <summary>
        /// Encapsulate the owner application to make accessible for child utils
        /// </summary>
        public new OutlookApi.Application OwnerApplication
        {
            get
            {
                return base.OwnerApplication as OutlookApi.Application;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Creates an instance of ApplicationUtils
        /// </summary>
        /// <returns>instance of ApplicationUtils</returns>
        protected internal virtual ApplicationUtils OnCreateApplicationUtils()
        {
            return new ApplicationUtils(this);
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Creates an instance of DialogUtils
        /// </summary>
        /// <returns>instance of DialogUtils</returns>
        protected override OfficeApi.Tools.Contribution.DialogUtils OnCreateDialogUtils()
        {
            return new OutlookDialogUtils(this);
        }

        /// <summary>
        /// Try to detect the visibility of host application main window.
        /// The implementation tries to find a visible Outlook application main window and returns true if
        /// it was found.
        /// </summary>
        /// <param name="defaultResult">fallback result if method fails</param>
        /// <returns>true if application is visible, otherwise false</returns>
        public override bool TryGetApplicationVisible(bool defaultResult)
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
