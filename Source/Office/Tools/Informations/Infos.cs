﻿using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Informations
{
    /// <summary>
    /// Contains various information subsets
    /// </summary>
    public class Infos
    {
        #region Fields

        private object _lock;
        private Contribution.CommonUtils _owner;
        private AssemblyInfo _assemblyInfo;
        private AppDomainInfo _appDomainInfo;
        private EnvironmentInfo _environmentInfo;
        private HostInfo _hostInfo;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        internal Infos(Contribution.CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _lock = new object();
            _owner = owner;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Executing Assembly Information
        /// </summary>
        public AssemblyInfo Assembly
        {
            get
            {
                lock (_lock)
                {
                    if (null == _assemblyInfo)
                        _assemblyInfo = _owner.OnCreateAssemblyInfo();                    
                }
                return _assemblyInfo;
            }
        }

        /// <summary>
        /// Current AppDomain Information
        /// </summary>
        public AppDomainInfo AppDomain
        {
            get
            {
                lock (_lock)
                {
                    if (null == _appDomainInfo)
                        _appDomainInfo = _owner.OnCreateAppDomainInfo();                    
                }
                return _appDomainInfo;
            }
        }

        /// <summary>
        /// Current Environment Information
        /// </summary>
        public EnvironmentInfo Environment
        {
            get
            {
                lock (_lock)
                {
                    if (null == _environmentInfo)
                        _environmentInfo = _owner.OnCreateEnvironmentInfo();                    
                }
                return _environmentInfo;
            }
        }

        /// <summary>
        /// Current Host Information
        /// </summary>
        public HostInfo Host
        {
            get
            {
                lock (_lock)
                {
                    if (null == _hostInfo)
                        _hostInfo = _owner.OnCreateHostInfo();                    
                }
                return _hostInfo;
            }
        }

        /// <summary>
        /// Owner Instance
        /// </summary>
        internal Contribution.CommonUtils Owner
        {
            get
            {
                return _owner;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Called from DiagnosticPairCollection to add custom system information
        /// </summary>
        /// <param name="diagnostics">sender instance</param>
        protected internal virtual void GetCustomInformations(DiagnosticPairCollection diagnostics)
        { 
        
        }

        #endregion
    }
}
