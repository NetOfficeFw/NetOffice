using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CoreServices;
using NetOffice.CollectionsGeneric;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// CoClass Application
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821758.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Access.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [HasInteropCompatibilityClass(typeof(ApplicationClass))]
    public class Application : _Application, NetOffice.AccessApi.Application, IAutomaticQuit, IApplicationVersionProvider
    {
		#pragma warning disable

		#region Fields

		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;

        private bool _versionRequested;
        private object _cachedVersion;
        private object _chachedVersionLock = new object();

        #endregion

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.AccessApi.Application);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>
        /// Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        /// <summary>
        /// Type Cache
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Application);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Application() : base()
		{

		}

        /// <summary>
        /// Creates a new instance of the class
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public Application(Core factory = null, bool tryProxyServiceFirst = false) : base()
        {
            object proxy = null;
            if (tryProxyServiceFirst)
            {
                proxy = ProxyService.GetActiveInstance("Access", "Application", false);
                if (null != proxy)
                {
                    CreateFromProxy(proxy, true);
                    FromProxyService = true;
                }
            }

            if(null == proxy)
            {
                CreateFromProgId("Access.Application", true);
            }

            Factory = null != factory ? factory : Core.Default;
            TryRequestVersion();
            RegisterAsApplicationVersionProvider();
            OnCreate();
            _isInitialized = true;
        }

        #endregion

        #region ICloneable<Application>

        /// <summary>
        /// Creates a new Application that is a copy of the current instance
        /// </summary>
        /// <returns>A new Application that is a copy of this instance</returns>
        /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public new virtual NetOffice.AccessApi.Application DeepCopy()
        {
            return base.Clone() as NetOffice.AccessApi.Application;
        }

        #endregion

        #region ICOMObjectProxyService

        /// <summary>
        /// Instance is created from an already running application
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual bool FromProxyService { get; private set; }

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running Access.Application instances from the environment/system
        /// </summary>
        /// <returns>Access.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return ProxyService.GetActiveInstances<Application>("Access", "Application");
        }

        /// <summary>
        /// Returns all running Access.Application instances from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>Access.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances(Func<Application, bool> predicate)
        {
            return ProxyService.GetActiveInstances<Application>("Access", "Application", predicate);
        }

        /// <summary>
        /// Returns the count of running Access.Application instances that passed a predicate filter
        /// </summary>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount()
        {
            var sequence = ProxyService.GetActiveInstances<Application>("Access", "Application");
            int result = sequence.Count;
            sequence.Dispose();
            return result;
        }

        /// <summary>
        /// Returns the count of running Access.Application instances that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount(Func<Application, bool> predicate)
        {
            var sequence = ProxyService.GetActiveInstances<Application>("Access", "Application", predicate);
            int result = sequence.Count;
            sequence.Dispose();
            return result;
        }

        /// <summary>
        /// Returns a running Access.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Access.Application instance or null(Nothing in Visual Basic)</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("Access", "Application", throwExceptionIfNotFound);
        }

        /// <summary>
        /// Returns first running Access.Application instance from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Access.Application instance or null(Nothing in Visual Basic)</returns>
        /// <exception cref="ArgumentOutOfRangeException">occurs if no instance match and throwExceptionIfNotFound is set</exception>
        public static Application GetActiveInstance(Func<Application, bool> predicate, bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("Access", "Application", predicate, throwExceptionIfNotFound);
        }

        #endregion

        #region IApplicationVersionProvider

        string IApplicationVersionProvider.Name
        {
            get
            {
                return "Microsoft Access";
            }
        }

        string IApplicationVersionProvider.ComponentName
        {
            get
            {
                return "NetOffice.AccessApi";
            }
        }

        /// <summary>
        /// Request version information on demand and cache to call the remote server only 1x times
        /// </summary>
        object IApplicationVersionProvider.Version
        {
            get
            {
                lock (_chachedVersionLock)
                {
                    if (null == _cachedVersion)
                    {
                        _cachedVersion = TryVersionPropertyGet();
                    }
                }
                return _cachedVersion;
            }
        }

        bool IApplicationVersionProvider.VersionRequested
        {
            get
            {
                return _versionRequested;
            }
        }

        void IApplicationVersionProvider.TryRequestVersion()
        {
            _cachedVersion = TryVersionPropertyGet();
        }

        /// <summary>
        /// Try get version information without fail
        /// </summary>
        /// <returns></returns>
        private object TryVersionPropertyGet()
        {
            try
            {
                if (null != _proxyShare)
                {
                    if (EntityIsAvailable("Version"))
                        return Invoker.PropertyGet(this, "Version");
                    else
                        return 9.0;
                }
                else
                    return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (null != _proxyShare)
                    _versionRequested = true;
            }
        }

        #endregion

        #region IAutomaticQuit

        /// <summary>
        /// Determines Quit method want be called while disposing if NetOffice.Settings.EnableAutomaticQuit is true.
        /// Default is true when instance has no parent object and its not a cloned instance, otherwise false.
        /// </summary>
        bool IAutomaticQuit.Enabled
        {

            get
            {
                return _callQuitInDispose;
            }
            set
            {
                _callQuitInDispose = value;
            }
        }

        #endregion

        #pragma warning restore
    }
}
