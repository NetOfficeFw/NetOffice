using System;
using System.Linq;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using NetOffice.Loader;
using NetOffice.CoreServices;
using NetOffice.CoreServices.Internal;
using NetOffice.Exceptions;
#if DEBUG
using NetOffice.Diagnostics.Internal;
#endif

namespace NetOffice
{
    #region IDispatch - imagine a world without...

    /// <summary>
    /// Exposes objects, methods and properties to programming tools and other applications that support Automation. COM components implement the IDispatch interface to enable access by Automation clients, such as Visual Basic.
    /// </summary>
    [Guid("00020400-0000-0000-c000-000000000046"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDispatch
    {
        /// <summary>
        /// Retrieves the number of type information interfaces that an object provides (either 0 or 1)
        /// </summary>
        /// <returns>
        /// This method can return one of these values
        /// S_OK - Success
        /// E_NOTIMPL - Failure
        /// </returns>
        [PreserveSig]
        int GetTypeInfoCount();

        /// <summary>
        /// Retrieves the type information for an object, which can then be used to get the type information for an interface
        /// </summary>
        /// <param name="iTInfo">The type information to return. Pass 0 to retrieve type information for the IDispatch implementation</param>
        /// <param name="lcid">The locale identifier for the type information. An object may be able to return different type information for different languages. This is important for classes that support localized member names. For classes that do not support localized member names, this parameter can be ignored</param>
        /// <returns>The requested type information object</returns>
        System.Runtime.InteropServices.ComTypes.ITypeInfo GetTypeInfo([MarshalAs(UnmanagedType.U4)] int iTInfo, [MarshalAs(UnmanagedType.U4)] int lcid);

        /// <summary>
        /// Maps a single member and an optional set of argument names to a corresponding set of integer DISPIDs, which can be used on subsequent calls to Invoke.
        /// </summary>
        /// <param name="riid">Reserved for future use. Must be IID_NULL</param>
        /// <param name="rgsNames">The array of names to be mapped</param>
        /// <param name="cNames">The count of the names to be mapped</param>
        /// <param name="lcid">The locale context in which to interpret the names</param>
        /// <param name="rgDispId">Caller-allocated array, each element of which contains an identifier (ID) corresponding to one of the names passed in the rgszNames array. The first element represents the member name. The subsequent elements represent each of the member's parameters</param>
        /// <returns>
        /// This method can return one of these values
        /// S_OK - Success
        /// E_OUTOFMEMORY - Out of memory
        /// DISP_E_UNKNOWNNAME - One or more of the specified names were not known. The returned array of DISPIDs contains DISPID_UNKNOWN for each entry that corresponds to an unknown name
        /// DISP_E_UNKNOWNLCID
        /// </returns> - The locale identifier (LCID) was not recognized
        [PreserveSig]
        int GetIDsOfNames(ref Guid riid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgsNames, int cNames, int lcid, [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);

        /// <summary>
        /// Provides access to properties and methods exposed by an object
        /// </summary>
        /// <param name="dispIdMember">Identifies the member. Use GetIDsOfNames or the object's documentation to obtain the dispatch identifier.</param>
        /// <param name="riid">Reserved for future use. Must be IID_NULL</param>
        /// <param name="lcid">The locale context in which to interpret arguments. The lcid is used by the GetIDsOfNames function, and is also passed to Invoke to allow the object to interpret its arguments specific to a locale</param>
        /// <param name="dwFlags">Flags describing the context of the Invoke call</param>
        /// <param name="pDispParams">Pointer to a DISPPARAMS structure containing an array of arguments, an array of argument DISPIDs for named arguments, and counts for the number of elements in the array</param>
        /// <param name="pVarResult">Pointer to the location where the result is to be stored, or NULL if the caller expects no result. This argument is ignored if DISPATCH_PROPERTYPUT or DISPATCH_PROPERTYPUTREF is specified</param>
        /// <param name="pExcepInfo">Pointer to a structure that contains exception information. This structure should be filled in if DISP_E_EXCEPTION is returned. Can be NULL</param>
        /// <param name="pArgErr">The index within rgvarg of the first argument that has an error. Arguments are stored in pDispParams->rgvarg in reverse order, so the first argument is the one with the highest index in the array. This parameter is returned only when the resulting return value is DISP_E_TYPEMISMATCH or DISP_E_PARAMNOTFOUND. This argument can be set to null</param>
        /// <returns>
        /// See http://msdn.microsoft.com/de-de/library/windows/desktop/ms221479%28v=vs.85%29.aspx
        /// </returns>
        [PreserveSig]
        int Invoke(int dispIdMember, ref Guid riid, [MarshalAs(UnmanagedType.U4)] int lcid, [MarshalAs(UnmanagedType.U4)] int dwFlags, ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out, MarshalAs(UnmanagedType.LPArray)] object[] pVarResult, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
    }

    #endregion

    /// <summary>
    /// Creation Factory for ICOMObject and derived types
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Core
    {
        #region Nested

        /// <summary>
        /// IsInitializedChanged delegate
        /// </summary>
        /// <param name="sender">sender instance</param>
        /// <param name="isInitialized">true if already initialized, otherwise false</param>
        public delegate void IsInitializedChangedHandler(Core sender, bool isInitialized);

        #endregion

        #region Fields

        /// <summary>
        /// the well known IUnknown Interface ID
        /// </summary>
        private static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        private static Core _default;
        private bool _initalized;
        private KnownKeyTokens _knownNetOfficeKeyTokens;
        private Assembly _thisAssembly;
        private Type _thisType;
        private static ICOMObject[] _emptyObjectSequence = new ICOMObject[0];
        private object _checkInitializeLock = new object();
        private object _thisAssemblyLock = new object();
        private object _assembliesLock = new object();
        private object _createComObjectLock = new object();
        private static object _defaultLock = new object();

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Core()
        {
            CoreDomain = new CurrentAppDomain(this);
            InternalObjectRegister = new CoreManagement(this);
            InternalObjectActivator = new CoreActivator(this);
            InternalObjectResolver = new CoreResolver(this);
            InternalObjectEvents = new CoreEvents(this);
            InternalFactories = new CoreFactories(this);
            InternalCache = new CoreCache(this);
            Console = OnCreateConsole();
            Settings = OnCreateSettings();
            Invoker = OnCreateInvoker();
            OnCreate();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="isDefault">mark this instance as default instance</param>
        protected internal Core(bool isDefault)
        {
            IsDefault = isDefault;
            CoreDomain = new CurrentAppDomain(this);
            InternalObjectRegister = new CoreManagement(this);
            InternalObjectActivator = new CoreActivator(this);
            InternalObjectResolver = new CoreResolver(this);
            InternalObjectEvents = new CoreEvents(this);
            InternalFactories = new CoreFactories(this);
            InternalCache = new CoreCache(this);
            Console = OnCreateConsole();
            Settings = OnCreateSettings();
            Invoker = OnCreateInvoker();
            OnCreate();
        }

        #endregion

        #region Events

        /// <summary>
        /// Occurs when the initialize state has been changed
        /// </summary>
        public event IsInitializedChangedHandler IsInitializedChanged;

        /// <summary>
        /// Raise the IsInitializedChanged event
        /// </summary>
        private void RaiseIsInitializedChanged()
        {
            var handler = IsInitializedChanged;
            if (null != handler)
                handler(this, _initalized);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Returns info about intialized state
        /// </summary>
        [Category("Core"), Description("The core is already initialized")]
        public bool IsInitialized
        {
            get
            {
                return _initalized;
            }
        }

        /// <summary>
        /// Shared Default Core
        /// </summary>
        public static Core Default
        {
            get
            {
                lock (_defaultLock)
                {
                    if (null == _default)
                        _default = new Core(true);
                    return _default;
                }
            }
        }

        /// <summary>
        /// Core Settings
        /// </summary>
        [Browsable(false)]
        public Settings Settings { get; internal set; }

        /// <summary>
        /// Core Console
        /// </summary>
        [Browsable(false)]
        public DebugConsole Console { get; internal set; }

        /// <summary>
        /// Core Invoker
        /// </summary>
        [Browsable(false)]
        public Invoker Invoker { get; internal set; }

        /// <summary>
        /// Returns information about the instance is the shared default core
        /// </summary>
        [Category("Core"), Description("The core is also the shared default core")]
        public bool IsDefault { get; private set; }

        /// <summary>
        /// Time that the initialize process has needed used to pass
        /// </summary>
        [Category("Core"), Description("Time that the initialize process has needed used to pass")]
        public TimeSpan InitializedTime { get; private set; }

        /// <summary>
        /// Cached instance type
        /// </summary>
        [Browsable(false)]
        public Type ThisType
        {
            get
            {
                if (null == _thisType)
                    _thisType = GetType();
                return _thisType;
            }
        }

        /// <summary>
        /// COM Object Register Services
        /// </summary>
        /// <remarks>Provides event based access to the NetOffice Instance Table</remarks>
        public ICoreManagement ObjectRegister
        {
            get
            {
                return InternalObjectRegister;
            }
        }

        /// <summary>
        /// COM Object Activation Services
        /// </summary>
        /// <remarks>Provides event based access to the NetOffice creation factory</remarks>
        [Category("Core"), Description("COM Object Activation Services")]
        public ICoreActivator ObjectActivator
        {
            get
            {
                return InternalObjectActivator;
            }
        }

        /// <summary>
        /// COM Object Resolve Services
        /// </summary>
        /// <remarks>Provides event based access to the NetOffice type resolver</remarks>
        [Category("Core"), Description("COM Object Resolve Services")]
        public ICoreResolver ObjectResolver
        {
            get
            {
                return InternalObjectResolver;
            }
        }

        /// <summary>
        /// COM Object Event Services
        /// </summary>
        /// <remarks>Provides event based access to the NetOffice Event Bridge system</remarks>
        [Category("Core"), Description("COM Object Event Services")]
        public ICoreEvents ObjectEvents
        {
            get
            {
                return InternalObjectEvents;
            }
        }

        /// <summary>
        /// Contains loaded factories
        /// </summary>
        /// <remarks>Provides access to loaded factories</remarks>
        [Category("Core"), Description("Factory Services")]
        public ICoreFactories Factories
        {
            get
            {
                return InternalFactories;
            }
        }

        /// <summary>
        /// Contains all cached types and entites
        /// </summary>
        /// <remarks>Provides access to cache system</remarks>
        [Category("Core"), Description("Cache Services")]
        public ICoreCache Cache
        {
            get
            {
                return InternalCache;
            }
        }

        /// <summary>
        /// Internal COM Object Event Register
        /// </summary>
        internal CoreManagement InternalObjectRegister { get; private set; }

        /// <summary>
        /// Internal COM Object Event Activator
        /// </summary>
        internal CoreActivator InternalObjectActivator { get; private set; }

        /// <summary>
        /// Internal COM Object Event Resolver
        /// </summary>
        internal CoreResolver InternalObjectResolver { get; private set; }

        /// <summary>
        /// Internal COM Object Event Services
        /// </summary>
        internal CoreEvents InternalObjectEvents { get; private set; }

        /// <summary>
        /// Internal Factory Holder
        /// </summary>
        internal CoreFactories InternalFactories { get; private set; }

        /// <summary>
        /// Internal type and entities cache
        /// </summary>
        internal CoreCache InternalCache { get; private set; }

        /// <summary>
        /// Current NetOffice Core Assembly
        /// </summary>
        internal Assembly ThisAssembly
        {
            get
            {
                lock (_thisAssemblyLock)
                {
                    if (null == _thisAssembly)
                        _thisAssembly = Assembly.GetAssembly(typeof(Core));
                }
                return _thisAssembly;
            }
        }

        /// <summary>
        /// Contains a list of all known netoffice assembly key tokens
        /// </summary>
        internal KnownKeyTokens KnownNetOfficeKeyTokens
        {
            get
            {
                if (null == _knownNetOfficeKeyTokens)
                    _knownNetOfficeKeyTokens = new KnownKeyTokensReader().Read();
                return _knownNetOfficeKeyTokens;
            }
        }

        /// <summary>
        /// Assembly Loader
        /// </summary>
        internal CurrentAppDomain CoreDomain { get; private set; }

        #endregion

        #region Methods

        /// <summary>
        /// Creates a new instance of Invoker
        /// </summary>
        /// <returns>invoker instance</returns>
        protected internal virtual Invoker OnCreateInvoker()
        {
            return IsDefault ? Invoker.Default : new Invoker(this);
        }

        /// <summary>
        /// Creates a new instance of Settings
        /// </summary>
        /// <returns>settings instance</returns>
        protected internal virtual Settings OnCreateSettings()
        {
            return IsDefault ? Settings.Default : new Settings();
        }

        /// <summary>
        /// Creates a new instance of DebugConsole
        /// </summary>
        /// <returns>DebugConsole instance</returns>
        protected internal virtual DebugConsole OnCreateConsole()
        {
            return IsDefault ? DebugConsole.Default : new DebugConsole();
        }

        /// <summary>
        /// Called from ctor at last
        /// </summary>
        protected internal virtual void OnCreate()
        {

        }

        #endregion

        #region Initialize

        /// <summary>
        /// Check for initialize state and call Initialize if its necessary
        /// </summary>
        /// <returns>initialize state</returns>
        /// <exception cref="NetOfficeInitializeException">unexpected error. see inner exception(s) for details.</exception>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public bool CheckInitialize()
        {
            lock (_checkInitializeLock)
            {
                if (!_initalized)
                {
#pragma warning disable 612, 618
                    Initialize();
#pragma warning restore 612, 618
                }
                return _initalized;
            }
        }

        /// <summary>
        /// Recieve factory infos from all loaded NetOfficeApi Assemblies in current application domain
        /// </summary>
        /// <exception cref="NetOfficeInitializeException">unexpected error. see inner exception(s) for details.</exception>
        [Obsolete("Not necessary anymore(self-initializing)")]
        public void Initialize()
        {
            Initialize(CacheOptions.KeepExistingCacheAlive);
        }

        /// <summary>
        /// Must be called from client assembly for ICOMObject Support
        /// Recieve factory infos from all loaded NetOfficeApi Assemblies in current application domain
        /// <param name="cacheOptions">NetOffice cache options</param>
        /// </summary>
        /// <exception cref="NetOfficeInitializeException">unexpected error. see inner exception(s) for details.</exception>
        [Obsolete("Not necessary anymore(self-initializing)")]
        public virtual void Initialize(CacheOptions cacheOptions)
        {
#if DEBUG
            new InternalDebugDiagnostics().ValidateCore(this);
#endif

            Settings.CacheOptions = cacheOptions;

            if (!_initalized)
            {
                _initalized = true;
                RaiseIsInitializedChanged();
            }

            try
            {
                DateTime startTime = DateTime.Now;

                lock (_assembliesLock)
                {
                    Console.WriteLine("NetOffice Core.Initialize() Core Version:{0}; Deep Loading:{1}; Load Assemblies Unsafe:{2}; AppDomain:{3}",
                         ThisAssembly.GetName().Version, Settings.EnableDeepLoading,
                         Settings.LoadAssembliesUnsafe, AppDomain.CurrentDomain.Id.ToString() + "-" + AppDomain.CurrentDomain.FriendlyName);

                    if (Settings.EnableMoreDebugOutput)
                    {
                        string localPath = Resolver.UriResolver.ResolveLocalPath(ThisAssembly.CodeBase);
                        Console.WriteLine("Local Bind Path:{0}", localPath);
                    }

                    CoreDomain.TryLoadAssemblies(this);

                    ClearCaches(false);
                    InternalFactories.LoadAPIFactories();
                    InternalFactories.LoadDependentAPIFactories();

                    InitializedTime = DateTime.Now - startTime;

                    if (Settings.EnableMoreDebugOutput)
                    {
                        Console.WriteLine("NetOffice Core contains {0} assemblies", InternalFactories.FactoryAssemblies.Count);
                        Console.WriteLine("NetOffice Core.Initialize() passed in {0} milliseconds", InitializedTime.TotalMilliseconds);
                    }
                }
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw new NetOfficeInitializeException(exception);
            }
        }

        /// <summary>
        /// Clears all Core caches
        /// </summary>
        /// <param name="forceClear">method want do nothing if cache option is KeepExistingCacheAlive. You can force clear caches anyway by giving true</param>
        public virtual void ClearCaches(bool forceClear)
        {
            if (forceClear || CacheOptions.ClearExistingCache == Settings.CacheOptions)
            {
                lock (_assembliesLock)
                {
                    InternalCache.Clear();
                    InternalFactories.Clear();
                    _initalized = false;
                }
            }
        }

        #endregion

        #region Create Unknown COMObject

        /// <summary>
        /// Creates a new ICOMObject based on classType of comProxy. The method use Settings.EnableDynamicEventArguments to reflect dynamics
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        /// <exception cref="ArgumentNullException">one ore more arguments is null</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        /// <exception cref="FactoryException">throws when its failed to find the corresponding factory. this indicates a missing netoffice api assembly</exception>
        /// <exception cref="NetOfficeInitializeException">unexpected initialization error. see inner exception(s) for details</exception>
        public virtual ICOMObject CreateEventArgumentObjectFromComProxy(ICOMObject caller, object comProxy)
        {
            return CreateObjectFromComProxy(caller, comProxy, caller.Settings.EnableDynamicEventArguments);
        }

        /// <summary>
        /// Creates a new ICOMObject array
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxyArray">new created proxy array</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>corresponding Wrapper class Instance array or plain COMObject array</returns>
        /// <exception cref="ArgumentNullException">one ore more arguments is null</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        /// <exception cref="FactoryException">throws when its failed to find the corresponding factory. this indicates a missing netoffice api assembly</exception>
        /// <exception cref="NetOfficeInitializeException">unexpected initialization error. see inner exception(s) for details</exception>
        public virtual ICOMObject[] CreateObjectArrayFromComProxy(ICOMObject caller, object[] comProxyArray, bool allowDynamicObject)
        {
            if (null == caller)
                throw new ArgumentNullException("caller");
            if (null == comProxyArray)
                throw new ArgumentNullException("comProxyArray");

            CheckInitialize();
            try
            {
                if (null == comProxyArray)
                    return null;

                lock (_createComObjectLock)
                {
                    Type comVariantType = null;
                    ICOMObject[] newVariantArray = new ICOMObject[comProxyArray.Length];
                    for (int i = 0; i < comProxyArray.Length; i++)
                    {
                        comVariantType = comProxyArray[i].GetType(); //  TODO: handle that better by a cache
                        newVariantArray[i] = CreateObjectFromComProxy(caller, comProxyArray[i], allowDynamicObject);
                    }
                    return newVariantArray;
                }
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
        }

        /// <summary>
        /// Creates a new ICOMObject based on classType of comProxy
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        /// <exception cref="ArgumentNullException">one ore more arguments is null</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        /// <exception cref="FactoryException">throws when its failed to find the corresponding factory. this indicates a missing netoffice api assembly</exception>
        /// <exception cref="NetOfficeInitializeException">unexpected initialization error. see inner exception(s) for details</exception>
        /// <exception cref="InvalidCastException">failed to convert result to type of t</exception>
        public virtual T CreateObjectFromComProxy<T>(ICOMObject caller, object comProxy, bool allowDynamicObject) where T :class, ICOMObject
        {
            object result = CreateObjectFromComProxy(caller, comProxy, allowDynamicObject);
            if (null != result)
                return default(T);
            else
                return (T)result;
        }

        /// <summary>
        /// Creates a new ICOMObject based on classType of comProxy
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        /// <exception cref="ArgumentNullException">comProxy arguments is null</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        /// <exception cref="FactoryException">throws when its failed to find the corresponding factory. this indicates a missing netoffice api assembly</exception>
        /// <exception cref="NetOfficeInitializeException">unexpected initialization error. see inner exception(s) for details</exception>
        public virtual ICOMObject CreateObjectFromComProxy(ICOMObject caller, object comProxy, bool allowDynamicObject)
        {
            if (null == caller)
                throw new ArgumentNullException("caller");

            CheckInitialize();
            try
            {
                ICOMObject result = null;
                Guid typeId = Guid.Empty;
                Guid componentId = Guid.Empty;

                CoreTypeExtensions.GetComponentAndTypeId(this, comProxy, ref componentId, ref typeId);

                lock (_createComObjectLock)
                {
                    // get type factory first to handle possible duplicate type
                    ITypeFactory factoryInfo = CoreFactoryExtensions.GetTypeFactory(this, caller, comProxy, componentId, typeId, false);
                    if (null != factoryInfo)
                    {
                        TypeInformation typeInfo = CoreTypeExtensions.GetTypeInformationForUnknownObject(this, factoryInfo, typeId, comProxy);
                        if(null == typeInfo)
                            throw new FactoryException(String.Format("Unable to resolve proxy type:{0}", ComTypes.TypeDescriptor.GetFullComponentClassName(comProxy)));

                        result = CoreCreateExtensions.CreateInstance(this, typeInfo, caller, comProxy);
                        //result = InternalObjectActivator.TryReplaceInstance(caller, result);
                    }
                    else
                    {
                        result = CoreCreateExtensions.TryCreateObjectByResolveEvent(this, caller, null, comProxy);
                        if (null == result)
                        {
                            if (allowDynamicObject && Settings.EnableDynamicObjects)
                            {
                                result = InternalObjectActivator.RaiseCreateCOMDynamic(caller, comProxy);
                                if (null == result)
                                    result = new COMDynamicObject(caller, comProxy);
                            }
                            else
                            {
                                throw new FactoryException(String.Format("Unable to resolve proxy type:{0}", ComTypes.TypeDescriptor.GetFullComponentClassName(comProxy)));
                            }
                        }
                    }

                    return result;
                }
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
        }

        #endregion

        #region Create Known COM Object

        /// <summary>
        /// Creates a new ICOMObject array based on wrapperClassType
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxyArray">new created proxies</param>
        /// <param name="contractWrapperType">type info from contract wrapper</param>
        /// <returns>corresponding wrapper class instances or plain COMObject</returns>
        /// <exception cref="ArgumentNullException">one ore more arguments is null</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        /// <exception cref="FactoryException">throws when its failed to find the corresponding factory. this indicates a missing netoffice api assembly</exception>
        /// <exception cref="NetOfficeInitializeException">unexpected initialization error. see inner exception(s) for details</exception>
        public virtual ICOMObject[] CreateKnownObjectArrayFromComProxy(ICOMObject caller, object[] comProxyArray, Type contractWrapperType)
        {
            CheckInitialize();
            try
            {
                if (null == comProxyArray)
                    return _emptyObjectSequence;

                ICOMObject[] newVariantArray = new ICOMObject[comProxyArray.Length];
                for (int i = 0; i < comProxyArray.Length; i++)
                {
                    ICOMObject newInstance = CreateKnownObjectFromComProxy(caller, comProxyArray[i], contractWrapperType);
                    newVariantArray[i] = newInstance;
                }
                return newVariantArray;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
        }

        /// <summary>
        /// Creates a new ICOMObject based on wrapperClassType
        /// </summary>
        /// <typeparam name="T">result contract type</typeparam>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="contractWrapperType">type info from contract wrapper</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        /// <exception cref="ArgumentNullException">comProxy or contractWrapperType is null</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        /// <exception cref="InvalidCastException">T is not equal or base from contractWrapperType argument</exception>
        /// <exception cref="FactoryException">throws when its failed to find the corresponding factory. this indicates a missing netoffice api assembly</exception>
        /// <exception cref="NetOfficeInitializeException">unexpected initialization error. see inner exception(s) for details</exception>
        public virtual T CreateKnownObjectFromComProxy<T>(ICOMObject caller, object comProxy, Type contractWrapperType) where T : class, ICOMObject
        {
            return (T)CreateKnownObjectFromComProxy(caller, comProxy, contractWrapperType);
        }

        /// <summary>
        /// Creates a new ICOMObject based on wrapperClassType
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="contractWrapperType">type info from contract wrapper</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        /// <exception cref="ArgumentNullException">comProxy or contractWrapperType</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        /// <exception cref="FactoryException">throws when its failed find to the corresponding factory. this indicates a missing netoffice api assembly</exception>
        /// <exception cref="NetOfficeInitializeException">unexpected initialization error. see inner exception(s) for details</exception>
        public virtual ICOMObject CreateKnownObjectFromComProxy(ICOMObject caller, object comProxy, Type contractWrapperType)
        {
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");
            if (null == contractWrapperType)
                throw new ArgumentNullException("contractWrapperType");

            if (Settings.EnableKnownReferenceInspection)
            {
                return CreateObjectFromComProxy(caller, comProxy, false);
            }
            else
            {
                return InternalCreateKnownObjectFromComProxy(caller, comProxy, contractWrapperType);
            }
        }

        private ICOMObject InternalCreateKnownObjectFromComProxy(ICOMObject caller, object comProxy, Type contractWrapperType)
        {
            CheckInitialize();
            try
            {
                ICOMObject result = null;
                lock (_createComObjectLock)
                {
                    TypeInformation typeInfo = CoreTypeExtensions.GetTypeInformationForKnownObject(this, contractWrapperType, comProxy);
                    if (null != typeInfo)
                    {
                        result = CoreCreateExtensions.CreateInstance(this, typeInfo, caller, comProxy);
                    }
                    else
                    {
                        result = CoreCreateExtensions.TryCreateObjectByResolveEvent(this, caller, contractWrapperType, contractWrapperType);
                        if (null == result)
                            throw new FactoryException(String.Format("Unable to find implementation: {0}.", contractWrapperType.FullName));
                    }

                    //result = InternalObjectActivator.TryReplaceInstance(caller, result);
                    return result;
                }
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
        }

        #endregion
    }
}
