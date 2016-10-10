using System;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Security.Principal;
using NetOffice.WordApi.Tools.Utils;

namespace NetOffice.WordApi.Tools.Utils
{
    /// <summary>
    /// Various helper for common tasks
    /// </summary>
    public class CommonUtils : NetOffice.OfficeApi.Tools.Utils.CommonUtils
    {
        #region Fields

        private WordApi.Application _ownerApplication;
        private FileUtils _fileUtils;
        private ApplicationUtils _applicationUtils;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the application
        /// </summary>
        /// <param name="application">owner application</param>
        public CommonUtils(WordApi.Application application): base(application)
        {
            _ownerApplication = application;
        }

        /// <summary>
        /// Creates an instance of the application
        /// </summary>
        /// <param name="application">owner application</param>
        /// <param name="ownerAssembly">owner assembly</param>
        public CommonUtils(WordApi.Application application, Assembly ownerAssembly) : base(application, ownerAssembly)
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
        /// File related utils
        /// </summary>
        public FileUtils File
        {
            get
            {
                if (null == _fileUtils)
                    _fileUtils = OnCreateFileUtils();
                return _fileUtils;
            }
        }

        /// <summary>
        /// Application related utils
        /// </summary>
        public ApplicationUtils Application
        {
            get 
            {
                if (null == _applicationUtils)
                    _applicationUtils = OnCreateApplicationUtils();
                return _applicationUtils;
            }
        }

        /// <summary>
        /// Encapsulate the owner application to make accessible for child utils
        /// </summary>
        public new WordApi.Application OwnerApplication
        {
            get
            {
                return base.OwnerApplication as WordApi.Application;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Creates an instance of FileUtils
        /// </summary>
        /// <returns>instance of FileUtils</returns>
        protected internal virtual FileUtils OnCreateFileUtils()
        {
            return new FileUtils(this);
        }

        /// <summary>
        /// Creates an instance of ApplicationUtils
        /// </summary>
        /// <returns>instance of ApplicationUtils</returns>
        protected internal virtual ApplicationUtils OnCreateApplicationUtils()
        {
            return new ApplicationUtils(this);
        }

        #endregion
    }
}
