using System;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Security.Principal;
using NetOffice.OfficeApi.Tools.Utils;
using NetOffice.OfficeApi.Tools.Informations;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// CommonUtils TryFindLoadLocation Result Value
    /// </summary>
    public enum RegistryLocationResult
    {
        /// <summary>
        /// Unable to find load key
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// Addin load key found in HKEY_LOCAL_MACHINE
        /// </summary>
        User = 1,

        /// <summary>
        /// Addin load key found in HKEY_CURRENT_USER
        /// </summary>
        System = 2
    }

    /// <summary>
    /// Various helper for common tasks
    /// </summary>
    public class CommonUtils : IDisposable
    {
        #region Fields

        private object _lock = new object();
        private const string _headerCaptionLineDefault = "----------";

        private string _headerCaptionLine;
        private Assembly _ownerAssembly;
        private ICOMObject _ownerApplication;
        private bool? _applicationIs2007OrHigher;
        private bool? _adminPermissions;
        private bool _isAutomation;
        private DialogUtils _dialogUtils;
        private ColorUtils _colorUtils;
        private ImageUtils _imageUtils;
        private TrayUtils _trayUtils;
        private ResourceUtils _resourceUtils;
        private Infos _infos;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the application
        /// </summary>
        /// <param name="application">owner application</param>
        public CommonUtils(ICOMObject application)
        {
            if (null == application)
                throw new ArgumentNullException("application");
            _ownerApplication = application;
            _headerCaptionLine = _headerCaptionLineDefault;
            _infos = new Infos(this);
        }

        /// <summary>
        /// Creates an instance of the application
        /// </summary>
        /// <param name="application">owner application</param>
        /// <param name="ownerAssembly">owner assembly</param>
        public CommonUtils(ICOMObject application, Assembly ownerAssembly)
        {
            if (null == application)
                throw new ArgumentNullException(nameof(application));
            if (null == ownerAssembly)
                throw new ArgumentNullException(nameof(ownerAssembly));
            _ownerApplication = application;
            _ownerAssembly = ownerAssembly;
            _headerCaptionLine = _headerCaptionLineDefault;
            _infos = new Infos(this);
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">addin owner instance</param>
        /// <param name="isAutomation">host application is started for automation</param>
        protected internal CommonUtils(NetOffice.Tools.COMAddinBase owner, bool isAutomation)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            Owner = owner;
            _ownerApplication = owner.AppInstance;
            _isAutomation = isAutomation;
            _headerCaptionLine = _headerCaptionLineDefault;
            _infos = new Infos(this);
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">addin owner</param>
        /// <param name="isAutomation">indicates the host application is currently in automation</param>
        /// <param name="ownerAssembly">owner application</param>
        protected internal CommonUtils(NetOffice.Tools.COMAddinBase owner, bool isAutomation, Assembly ownerAssembly)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            if (null == ownerAssembly)
                throw new ArgumentNullException("ownerAssembly");
            Owner = owner;
            _ownerApplication = owner.AppInstance;
            _ownerAssembly = ownerAssembly;
            _isAutomation = isAutomation;
            _headerCaptionLine = _headerCaptionLineDefault;
            _infos = new Infos(this);
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">addin owner</param>
        /// <param name="ownerType">type information from addin owner</param>
        /// <param name="isAutomation">indicates the host application is currently in automation</param>
        /// <param name="ownerAssembly">owner application</param>
        protected internal CommonUtils(NetOffice.Tools.COMAddinBase owner, Type ownerType, bool isAutomation, Assembly ownerAssembly)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            if (null == ownerAssembly)
                throw new ArgumentNullException("ownerAssembly");
            Owner = owner;
            _ownerApplication = owner.AppInstance;
            _ownerAssembly = ownerAssembly;
            _isAutomation = isAutomation;
            _headerCaptionLine = _headerCaptionLineDefault;
            _infos = new Infos(this);
        }

        #endregion
          
        #region Properties

        /// <summary>
        /// Addin Owner Instance. Can be null if its used in custom application
        /// </summary>
        public NetOffice.Tools.COMAddinBase Owner { get; private set; }

        /// <summary>
        /// Cached Owner Type Information
        /// </summary>
        private Type OwnerType { get; set; }

        /// <summary>
        /// Host/Office Application
        /// </summary>
        public ICOMObject OwnerApplication
        {
            get
            {
                return _ownerApplication;
            }
        }

        /// <summary>
        /// The Office Host Application is in Version 12.00 or higher
        /// </summary>
        public bool ApplicationIs2007OrHigher
        {
            get
            {
                lock (_lock)
                {
                    if (null == _applicationIs2007OrHigher)
                    {
                        double? version = TryGetApplicationVersion();
                        if (null != version && version >= 12.00)
                            _applicationIs2007OrHigher = true;
                        else
                            _applicationIs2007OrHigher = false;
                    }                    
                }
                return (bool)_applicationIs2007OrHigher;
            }
        }

        /// <summary>
        /// Current domain has elevated permissions
        /// </summary>
        public bool AdminPermissions
        {
            get
            {
                lock (_lock)
                {
                    if (null == _adminPermissions)
                    {
                        WindowsIdentity identity = WindowsIdentity.GetCurrent();
                        WindowsPrincipal principal = new WindowsPrincipal(identity);
                        bool result = principal.IsInRole(WindowsBuiltInRole.Administrator);
                        identity.Dispose();
                        _adminPermissions = result;
                    }   
                }
                return (bool)_adminPermissions;
            }
        }

        /// <summary>
        /// The host application is currently in automation mode. In this case, avoid any kind of dialogs or something like that 
        /// </summary>
        public bool IsAutomation
        {
            get
            {
                return _isAutomation;
            }
        }

        /// <summary>
        /// Dialog related utils
        /// </summary>
        public DialogUtils Dialog
        {
            get
            {
                lock (_lock)
                {
                    if (null == _dialogUtils)
                        _dialogUtils = OnCreateDialogUtils();
                }
                return _dialogUtils;
            }
        }

        /// <summary>
        /// Resource related utils
        /// </summary>
        public ResourceUtils Resource
        {
            get
            {
                lock (_lock)
                {
                    if (null == _resourceUtils)
                        _resourceUtils = OnCreateResourceUtils();
                }
                return _resourceUtils;
             }
        }

        /// <summary>
        /// Tray related utils
        /// </summary>
        public TrayUtils Tray
        {
            get
            {
                lock (_lock)
                {
                    if (null == _trayUtils)
                        _trayUtils = OnCreateTrayUtils();                    
                }
                return _trayUtils;
            }
        }

        /// <summary>
        /// Image related utils
        /// </summary>
        public ImageUtils Image
        {
            get
            {
                lock (_lock)
                {
                    if (null == _imageUtils)
                        _imageUtils = OnCreateImageUtils();                    
                }
                return _imageUtils;
            }
        }

        /// <summary>
        /// Color related utils
        /// </summary>
        public ColorUtils Color
        {
            get
            {
                lock (_lock)
                {
                    if (null == _colorUtils)
                        _colorUtils = OnCreateColorUtils();                    
                }
                return _colorUtils;
            }
        }

        /// <summary>
        /// Various system informations
        /// </summary>
        public Infos Infos
        {
            get
            {
                lock (_lock)
                {
                    if (null == _infos)
                        _infos = new Infos(this);                    
                }
                return _infos;
            }
        }

        /// <summary>
        /// Fill header line in summary info as visual seperator
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public string HeaderCaptionLine
        {
            get
            {
                return _headerCaptionLine;
            }
            set
            {
                _headerCaptionLine = value;
            }
        }

        /// <summary>
        /// Header line for summary if ToolsUtils is used in custom applications
        /// </summary>
        internal static string HeaderCaptionLineDefault
        {
            get
            {
                return _headerCaptionLineDefault;
            }
        }

        /// <summary>
        /// Assembly informations used in AssemblyInfo
        /// </summary>
        protected internal Assembly OwnerAssembly
        {
            get
            {
                lock (_lock)
                {
                    if (null == _ownerAssembly)
                        _ownerAssembly = Assembly.GetExecutingAssembly();
                }
                return _ownerAssembly;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Try to detect the registry load location from an addin
        /// </summary>
        /// <param name="addinType">addin class type informations</param>
        /// <param name="applicationType">which office application</param>
        /// <returns>location or unkown</returns>
        public static RegistryLocationResult TryFindAddinLoadLocation(Type addinType, ApplicationIdentifiers.ApplicationType applicationType)
        {
            try
            {
                System.Runtime.InteropServices.ProgIdAttribute progId =
                  NetOffice.Tools.AttributeReflector.GetProgIDAttribute(addinType, false);
                if (null == progId)
                    return RegistryLocationResult.Unknown;

                string path = String.Format("Software\\Microsoft\\Office\\{0}\\Addins",
                    ApplicationIdentifiers.ConvertApplicationType(applicationType));

                Microsoft.Win32.RegistryKey key =
                    Microsoft.Win32.Registry.CurrentUser.OpenSubKey(path, false);
                if (null != key)
                {
                    key.Close();
                    return RegistryLocationResult.User;
                }

                try
                {
                    key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(path, false);
                    if (null != key)
                    {
                        key.Close();
                        return RegistryLocationResult.System;
                    }
                }
                catch (System.Security.SecurityException)
                {
                    return RegistryLocationResult.Unknown;
                }
                catch
                {
                    throw;
                }

                return RegistryLocationResult.Unknown;
            }
            catch (Exception)
            {
                return RegistryLocationResult.Unknown;
            }
        }

        /// <summary>
        /// Try to detect the registry load location from an addin
        /// </summary>
        /// <returns>load location or information it is unable to detect</returns>
        public RegistryLocationResult TryFindAddinLoadLocation()
        {
            try
            {
                Type addinType = null;
                if (null != OwnerType)
                    addinType = OwnerType;
                else if (null != Owner)
                    addinType = Owner.GetType();
                else
                    return RegistryLocationResult.Unknown;

                System.Runtime.InteropServices.ProgIdAttribute progId =
                  NetOffice.Tools.AttributeReflector.GetProgIDAttribute(addinType, false);
                if (null == progId)
                    return RegistryLocationResult.Unknown;

                ApplicationIdentifiers.ApplicationType applicationType =
                    ApplicationIdentifiers.IsApplication(OwnerApplication.UnderlyingType.GUID);
                if (ApplicationIdentifiers.ApplicationType.None == applicationType)
                    return RegistryLocationResult.Unknown;

                return CommonUtils.TryFindAddinLoadLocation(addinType, applicationType);
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Creates an instance of DialogUtils
        /// </summary>
        /// <returns>instance of DialogUtils</returns>
        protected internal virtual DialogUtils OnCreateDialogUtils()
        {
            return new DialogUtils(this);
        }

        /// <summary>
        /// Creates an instances of ResourceUtils
        /// </summary>
        /// <returns>instance of ResourceUtils</returns>
        protected internal virtual ResourceUtils OnCreateResourceUtils()
        {
            return new ResourceUtils(this);
        }

        /// <summary>
        /// Creates an instance of TrayUtils
        /// </summary>
        /// <returns>instance of TrayUtils</returns>
        protected internal virtual TrayUtils OnCreateTrayUtils()
        {
            return new TrayUtils(this);
        }

        /// <summary>
        /// Creates an instance of ImageUtils
        /// </summary>
        /// <returns>instance of ImageUtils</returns>
        protected internal virtual ImageUtils OnCreateImageUtils()
        {
            return new ImageUtils(this);        
        }

        /// <summary>
        /// Creates an instance of FileUtils
        /// </summary>
        /// <returns>instance of ColorUtils</returns>
        protected internal virtual ColorUtils OnCreateColorUtils()
        {
            return new ColorUtils(this);
        }

        /// <summary>
        /// Creates an instance of AssemblyInfo
        /// </summary>
        /// <returns>instance of AssemblyInfo</returns>
        protected internal virtual AssemblyInfo OnCreateAssemblyInfo()
        {
            return new AssemblyInfo(this);
        }

        /// <summary>
        /// Creates an instance of AppDomainInfo
        /// </summary>
        /// <returns>instance of AppDomainInfo</returns>
        protected internal virtual AppDomainInfo OnCreateAppDomainInfo()
        {
            return new AppDomainInfo(this);
        }

        /// <summary>
        /// Creates an instance of EnvironmentInfo
        /// </summary>
        /// <returns>instance of EnvironmentInfo</returns>
        protected internal virtual EnvironmentInfo OnCreateEnvironmentInfo()
        {
            return new EnvironmentInfo(this);
        }

        /// <summary>
        /// Creates an instance of HostInfo
        /// </summary>
        /// <returns>instance of HostInfo</returns>
        protected internal virtual HostInfo OnCreateHostInfo()
        {
            return new HostInfo(this);
        }

        /// <summary>
        /// Try to detect current host application version. (All MS-Office products supports the version property except for Access9 aka Access 2000)
        /// </summary>
        /// <returns>version or null if its failed to detect</returns>
        protected internal double? TryGetApplicationVersion()
        {
            try
            {
                if (_ownerApplication.EntityIsAvailable("Version"))
                {
                    double version = Convert.ToDouble(_ownerApplication.Invoker.PropertyGet(_ownerApplication, "Version"), CultureInfo.InvariantCulture);
                    return version;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null; 
            }
        }

        #endregion

        #region IDisposable

        /// <summary>
        /// Dispose the instance and cleanup/discard resources
        /// </summary>
        public virtual void Dispose()
        {
            if (null != _trayUtils)
                _trayUtils.DisposeTray();
        }

        #endregion
    }
}
