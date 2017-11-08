using System;
using System.Linq;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using NetOffice.Loader;
using NetOffice.Duck;
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
    public partial class Core
    {
        #region Fields
        
        /// <summary>
        /// the well know IUnknown Interface ID
        /// </summary>
        private static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        private Dictionary<Type, Type> _duckingCache;
        private static Core _default;
        private bool _initalized;
        private List<ICOMObject> _globalObjectList = new List<ICOMObject>();
        private Dictionary<string, Type> _proxyTypeCache = new Dictionary<string, Type>();
        private Dictionary<string, Type> _wrapperTypeCache = new Dictionary<string, Type>();
        private KnownKeyTokens _knownNetOfficeKeyTokens;
        private Assembly _thisAssembly;
        private static Type _thisType;
        private static ICOMObject[] _emptyOwnerPath = new ICOMObject[0];

        private object _checkInitializeLock = new object();
        private object _thisAssemblyLock = new object();        
        private object _factoryListLock = new object();
        private object _comObjectLock = new object();
        private static object _defaultLock = new object();

        #endregion
        
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Core()
        {
            Assemblies = new FactoryList();
            DependentAssemblies = new List<DependentAssembly>();
            CoreDomain = new CurrentAppDomain(this);
            Settings = new Settings();
            Console = new DebugConsole();
            Invoker = new Invoker(this);
            EntitiesListCache = new Dictionary<string, Dictionary<string, string>>();
            HostCache = new Dictionary<Guid, Guid>();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="isDefault">mark this instance as default instance</param>
        private Core(bool isDefault)
        {
            Assemblies = new FactoryList();
            DependentAssemblies = new List<DependentAssembly>();
            CoreDomain = new CurrentAppDomain(this);
            EntitiesListCache = new Dictionary<string, Dictionary<string, string>>();
            HostCache = new Dictionary<Guid, Guid>();
            IsDefault = isDefault;
            if (IsDefault)
            {
                Settings = Settings.Default;
                Console = DebugConsole.Default;
                Invoker = Invoker.Default;
            }
            else
            {
                Settings = new Settings();
                Console = new DebugConsole();
                Invoker = new Invoker(this);
            }
        }

        #endregion

        #region Events

        /// <summary>
        /// Occurs when its failed to resolve a wrapper type for a given com proxy
        /// </summary>
        public event ResolveEventHandler Resolve;

        /// <summary>
        /// Raise Resolve event
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="fullClassName">target NetOffice class</param>
        /// <param name="comProxy">native proxy type</param>
        /// <returns>type to use or null</returns>
        private Type RaiseResolve(ICOMObject caller, string fullClassName, Type comProxy)
        {
            if (null != Resolve)
            {
                ResolveEventArgs args = new ResolveEventArgs(caller, fullClassName, comProxy);
                Resolve(this, args);
                return args.Result;
            }
            else
                return null;
        }

        /// <summary>
        /// Occurs when a new COMDynamicObject instance should be created
        /// </summary>
        public event OnCreateCOMDynamicEventHandler CreateCOMDynamic;

        /// <summary>
        /// Raise the CreateCOMDynamic event
        /// </summary>
        /// <param name="instance">requested instance</param>
        /// <param name="comProxy">target proxy</param>
        /// <returns>COMDynamicObject instance or null</returns>
        private COMDynamicObject RaiseCreateCOMDynamic(ICOMObject instance, object comProxy)
        {
            if (null != CreateCOMDynamic)
            {
                OnCreateCOMDynamicEventArgs args = new OnCreateCOMDynamicEventArgs(instance, comProxy);
                CreateCOMDynamic(this, args);
                return args.Result;
            }
            else
                return null;
        }

        /// <summary>
        /// Occurs when a new COMProxyShare instance should be created
        /// </summary>
        public event OnCreateProxyShareEventHandler CreateProxyShare;

        /// <summary>
        /// Raise the CreateProxyShare event
        /// </summary>
        /// <param name="instance">requested instance</param>
        /// <param name="isEnumerator">indicates rcw is an enumerator</param>
        /// <returns>CreateProxyShare instance or null</returns>
        private COMProxyShare RaiseCreateProxyShare(ICOMObject instance, bool isEnumerator)
        {
            if (null != CreateProxyShare)
            {
                OnCreateProxyShareEventArgs args = new OnCreateProxyShareEventArgs(instance, isEnumerator);
                CreateProxyShare(this, args);
                return args.Result;
            }
            else
                return null;
        }

        /// <summary>
        /// Occours when a new COMObject instance has been created
        /// </summary>
        public event OnCreateInstanceEventHandler CreateInstance;

        /// <summary>
        /// Raise CreateInstance event
        /// </summary>
        /// <param name="instance">origin instance</param>
        /// <param name="replace">type to replace the instance</param>
        private void RaiseCreateInstance(ICOMObject instance, ref Type replace)
        {
            if (null != CreateInstance)
            {
                OnCreateInstanceEventArgs args = new OnCreateInstanceEventArgs(instance);
                CreateInstance(this, args);
                replace = args.Replace;
            }
        }

        /// <summary>
        /// Occurs when the initialize state has been changed
        /// </summary>
        public event IsInitializedChangedHandler IsInitializedChanged;

        /// <summary>
        /// Raise the IsInitializedChanged event
        /// </summary>
        private void RaiseIsInitializedChanged()
        {
            if (null != IsInitializedChanged)
                IsInitializedChanged(_initalized);
        }

        /// <summary>
        /// Notify info the count of proxies there open are changed
        /// in case of notify comes from event trigger created proxy the call comes from other thread
        /// </summary>
        public event ProxyCountChangedHandler ProxyCountChanged;

        /// <summary>
        /// Raise the ProxyCountChanged event (and optional, send channel message to console)
        /// </summary>
        /// <param name="proxyCount">current count of open com proxies</param>
        private void RaiseProxyCountChanged(int proxyCount)
        {
            try
            {
                if (null != ProxyCountChanged)
                    ProxyCountChanged(proxyCount);
            }
            catch (Exception throwedException)
            {
                DebugConsole.Default.WriteException(throwedException);
            }
        }

        /// <summary>
        /// Occurs when a proxy has been added
        /// </summary>
        public event ProxyAddedHandler ProxyAdded;

        /// <summary>
        /// ProxyAdded event has currently recipients
        /// </summary>
        internal bool HasProxyAddedRecipients
        {
            get
            {
                return (null != ProxyAdded);
            }
        }

        private void RaiseProxyAdded(IEnumerable<ICOMObject> ownerPath, ICOMObject item)
        {
            if (null != ProxyAdded)
                ProxyAdded(this, ownerPath, item);
        }

        /// <summary>
        ///  Occurs when a proxy has been removed
        /// </summary>
        public event ProxyRemovedHandler ProxyRemoved;

        /// <summary>
        /// ProxyRemoved event has currently recipients
        /// </summary>
        internal bool HasProxyRemovedRecipients
        {
            get
            {
                return (null != ProxyRemoved);
            }
        }

        private void RaiseProxyRemoved(IEnumerable<ICOMObject> ownerPath, ICOMObject item)
        {
            if (null != ProxyRemoved)
                ProxyRemoved(this, ownerPath, item);
        }

        /// <summary>
        /// Occurs when all proxies has been removed
        /// </summary>
        public event ProxyClearHandler ProxyCleared;

        private void RaiseProxyCleared()
        {
            if (null != ProxyCleared)
                ProxyCleared(this);
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
        /// Returns a sequence of currently loaded NetOffice API assemblies
        /// </summary>
        [Browsable(false)]
        public FactoryList Assemblies { get; private set; }
        
        /// <summary>
        /// Returns current count of open proxies
        /// </summary>
        [Category("Core"), Description("Current count of open proxies")]
        public int ProxyCount
        {
            get
            {
                return _globalObjectList.Count;
            }
        }

        /// <summary>
        /// Time that the initialize process has been used to pass
        /// </summary>
        [Category("Core"), Description("Time that the initialize process has been used to pass")]
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
        /// Contains a list of all known netoffice assembly key tokens
        /// </summary>
        [Browsable(false)]
        public KnownKeyTokens KnownNetOfficeKeyTokens
        {
            get
            {
                if (null == _knownNetOfficeKeyTokens)
                {
                    string[] tokens = CurrentAppDomain.KeyTokens(this);
                    _knownNetOfficeKeyTokens = new KnownKeyTokens();
                    foreach (string item in tokens)
                        _knownNetOfficeKeyTokens.Add(item);
                }
                return _knownNetOfficeKeyTokens;
            }
        }

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
                        _thisAssembly = Assembly.GetAssembly(typeof(COMObject));
                }
                return _thisAssembly;
            }
        }
        
        /// <summary>
        /// Assembly Loader
        /// </summary>
        internal CurrentAppDomain CoreDomain { get; private set; }

        /// <summary>
        /// ICOMObjectAvaility Cache
        /// </summary>
        internal Dictionary<string, Dictionary<string, string>> EntitiesListCache { get; private set; }

        /// <summary>
        /// Cache as Type ID => ParentLibrary ID
        /// </summary>
        internal Dictionary<Guid, Guid> HostCache { get; private set; }
        
        /// <summary>
        /// Duck Type Cache
        /// T1 is interface
        /// T2 is its implementation
        /// </summary>
        private Dictionary<Type, Type> DuckingCache
        {
            get
            {
                lock (_thisAssemblyLock)
                {
                    if (null == _duckingCache)
                        _duckingCache = new Dictionary<Type, Type>();
                }
                return _duckingCache;
            }
        }

        /// <summary>
        /// Dependent assemblies analyzed by LoadAPIFactories
        /// </summary>
        private List<DependentAssembly> DependentAssemblies { get; set; }

        #endregion
        
        #region Factory Methods

        /// <summary>
        /// Recieve factory infos from all loaded NetOfficeApi Assemblies in current application domain
        /// </summary>
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
        [Obsolete("Not necessary anymore(self-initializing)")]
        public void Initialize(CacheOptions cacheOptions)
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

                lock (_factoryListLock)
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
                    LoadAPIFactories();
                    LoadDependentAPIFactories();

                    InitializedTime = DateTime.Now - startTime;

                    if (Settings.EnableMoreDebugOutput)
                    {
                        Console.WriteLine("NetOffice Core contains {0} assemblies", Assemblies.Count);
                        Console.WriteLine("NetOffice Core.Initialize() passed in {0} milliseconds", InitializedTime.TotalMilliseconds);
                    }
                }            
            }
            catch (Exception throwedException)
            {
                DebugConsole.Default.WriteException(throwedException);
                throw;
            }
        }

        /// <summary>
        /// Check for initialize state and call Initialize if its necessary
        /// </summary>
        /// <returns>initialize state</returns>
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
        /// Clears all Core caches
        /// </summary>
        /// <param name="forceClear">method want do nothing if cache option is KeepExistingCacheAlive. You can force clear caches anyway by giving true</param>
        public void ClearCaches(bool forceClear)
        {
            if (forceClear || CacheOptions.ClearExistingCache == Settings.CacheOptions)
            {
                _wrapperTypeCache.Clear();
                HostCache.Clear();
                _proxyTypeCache.Clear();
                EntitiesListCache.Clear();
                Assemblies.Clear();
                DependentAssemblies.Clear();
            }
        }

        /// <summary>
        /// Get wrapper class factory info as non duck
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="throwException">throw exception if no info found or return null</param>
        /// <returns>factory info from corresponding assembly</returns>
        internal IFactoryInfo GetInstanceFactoryInfo(ICOMObject caller, object comProxy, bool throwException = true)
        {
            return this.GetFactoryInfo(HostCache, caller, comProxy, false, throwException);
        }

        /// <summary>
        ///  Get wrapper class factory info as duck
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="throwException">throw exception if no info found or return null</param>
        /// <returns>factory info from corresponding assembly</returns>
        internal IFactoryInfo GetDuckFactoryInfo(ICOMObject caller, object comProxy, bool throwException = true)
        {
            return this.GetFactoryInfo(HostCache, caller, comProxy, true, throwException);
        }

        /// <summary>
        /// Analyze assemblies in current appdomain and connect all NetOffice API factories to the core runtime.
        /// </summary>
        private void LoadAPIFactories()
        {
            DependentAssemblies.Clear();
            Assembly[] assemblies = CoreDomain.GetAssemblies();
            foreach (Assembly itemAssembly in assemblies)
            {
                string assemblyName = itemAssembly.GetName().Name;
                if (KnownNetOfficeKeyTokens.ContainsNetOfficeAttribute(itemAssembly))
                {                    
                    string[] depends = RecieveAssemblyFactory(assemblyName, itemAssembly);
                    foreach (string depend in depends)
                    {
                        if (!DependentAssemblies.Any(e => e.Name == depend))
                            DependentAssemblies.Add(new DependentAssembly(depend, itemAssembly));
                    }
                }

                if (Settings.EnableDeepLoading)
                {
                    foreach (AssemblyName itemName in itemAssembly.GetReferencedAssemblies())
                    {
                        if (KnownNetOfficeKeyTokens.ContainsNetOfficePublicKeyToken(itemName))
                        {
                            Assembly deepAssembly = CoreDomain.Load(itemName);
                            if (null == deepAssembly)
                                continue;

                            string deepAssemblyName = itemName.Name;
                            string[] depends = RecieveAssemblyFactory(deepAssemblyName, deepAssembly);
                            foreach (string depend in depends)
                            {
                                if (!DependentAssemblies.Any(e => e.Name == depend))
                                    DependentAssemblies.Add(new DependentAssembly(depend, itemAssembly));
                            }
                        }
                    }
                }
            }
        }
        
        /// <summary>
        /// Analyze dependent assemblies and connect there NetOffice API factories to the core runtime
        /// </summary>
        private void LoadDependentAPIFactories()
        {
            if (!Settings.EnableAdHocLoading)
                return;

            foreach (DependentAssembly dependAssembly in DependentAssemblies)
            {
                if (!Assemblies.Contains(dependAssembly.Name))
                {
                    string fileName = PathBuilder.BuildLocalPathFromDependentAssembly(dependAssembly);                   
                    if (System.IO.File.Exists(fileName))
                    {
                        try
                        {
                            Assembly asssembly = CoreDomain.Load(fileName);
                            RecieveAssemblyFactory(asssembly.GetName().Name, asssembly);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteException(exception);
                        }
                    }
                    else
                    {
                        Console.WriteLine(string.Format("Assembly {0} not found.", fileName));
                    }
                }
            }
        }

        /// <summary>
        /// Recieve factory instance from assembly and add them to factory cache
        /// </summary>
        /// <param name="name">name of the assembly</param>
        /// <param name="assembly">assemmbly to recieve</param>
        /// <returns>array of dependend assemblies</returns>
        private string[] RecieveAssemblyFactory(string name, Assembly assembly)
        {
            if (false == Attributes.NetOfficeAssemblyAttribute.ContainsAttribute(assembly))
                return new string[0];
            
            NetOffice.IFactoryInfo factoryInfo = Assemblies.FirstOrDefault(e => e.AssemblyName == name);
            if (null == factoryInfo)
            {
                List<string> dependAssemblies = new List<string>();
                Type factoryInfoType = assembly.GetType(name + ".Utils.ProjectInfo");
                if (null == factoryInfoType)
                    throw new NetOfficeException(String.Format("Unable to find {0} factory info", name));
                factoryInfo = Activator.CreateInstance(factoryInfoType) as IFactoryInfo;
                if (null == factoryInfo)
                    throw new FactoryException(String.Format("Unexpected {0} factory info. Assembly {0}", name, assembly));
                Assemblies.Add(factoryInfo);
                Console.WriteLine("NetOffice Core recieved IFactoryInfo:{0}:{1}", factoryInfo.Assembly.FullName, factoryInfo.Assembly.FullName);

                foreach (string itemDependency in factoryInfo.Dependencies)
                    dependAssemblies.Add(itemDependency);

                return dependAssemblies.ToArray();
            }
            else
                return new string[0];
        }

        #endregion

        #region Create COMObject Methods

        /// <summary>
        /// Creates a new COMProxyShare instance
        /// </summary>
        /// <param name="sender">requested instance</param>
        /// <param name="comProxy">inner proxy rcw</param>
        ///  <param name="isEnumerator">indicates rcw is an enumerator</param>
        /// <returns>new instance</returns>
        /// <exception cref="CreateCOMProxyShareException">throws when its failed to create instance</exception>
        public COMProxyShare CreateNewProxyShare(ICOMObject sender, object comProxy, bool isEnumerator)
        {
            try
            {
                COMProxyShare instance = RaiseCreateProxyShare(sender, isEnumerator);
                return null != instance ? instance : new COMProxyShare(this, comProxy, isEnumerator);
            }
            catch (Exception exception)
            {
                throw new CreateCOMProxyShareException(exception);
            }          
        }

        /// <summary>
        /// Creates a new COMProxyShare instance
        /// </summary>
        /// <param name="sender">requested instance</param>
        /// <param name="comProxy">inner proxy rcw</param>
        /// <returns>new instance</returns>
        /// <exception cref="CreateCOMProxyShareException">throws when its failed to create instance</exception>
        public COMProxyShare CreateNewProxyShare(ICOMObject sender, object comProxy)
        {
            try
            {
                COMProxyShare instance = RaiseCreateProxyShare(sender, false);
                return null != instance ? instance : new COMProxyShare(this, comProxy);
            }
            catch (Exception exception)
            {
                throw new CreateCOMProxyShareException(exception);
            }            
        }

        /// <summary>
        /// Creates a new duck typing instance by given generic type argument
        /// </summary>    
        /// <typeparam name="T">interface result type</typeparam>
        /// <returns>new instance</returns>
        /// <exception cref="ArgumentException">throws when ComProgIdAttribute is missing</exception>
        /// <exception cref="DuckException">throws when its failed to compile an implementation</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception> 
        /// <exception cref="COMException">throws when its failed to recieve progID Type</exception> 
        public T CreateDuckObject<T>() where T : ICOMObject
        {
            object[] attributes = typeof(T).GetCustomAttributes(typeof(NetOffice.Attributes.ComProgIdAttribute), false);
            if (attributes.Length > 0)
            {
                NetOffice.Attributes.ComProgIdAttribute attribute = attributes[0] as NetOffice.Attributes.ComProgIdAttribute;
                return CreateDuckObject<T>(attribute.Value);
            }
            else
                throw new ArgumentException("ComProgIdAttribute is missing.");
        }

        /// <summary>
        /// Creates a new duck typing instance by given generic type argument
        /// </summary>
        /// <typeparam name="T">interface result type</typeparam>
        /// <param name="progId">progId to create</param>
        /// <returns>new instance</returns>
        /// <exception cref="ArgumentNullException">throws when progId is null or empty</exception>
        /// <exception cref="DuckException">throws when its failed to compile an implementation</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception> 
        /// <exception cref="COMException">throws when its failed to recieve progID Type</exception> 
        public T CreateDuckObject<T>(string progId) where T : ICOMObject
        {
            if (String.IsNullOrWhiteSpace(progId))
                throw new ArgumentNullException("progId");
            Type type = System.Type.GetTypeFromProgID(progId, false);
            if (null == type)
                throw new COMException("Unable to recieve progId Type:<" + progId + ">");

            object interopProxy = null;

            try
            {
                interopProxy = Activator.CreateInstance(type);
            }
            catch (Exception exception)
            {
                throw new CreateInstanceException(exception);
            }
             
            return CreateDuckObjectFromComProxy<T>(interopProxy);
        }

        /// <summary>
        /// Creates a new duck typing instance by given generic type argument
        /// </summary>
        /// <typeparam name="T">interface result type</typeparam>
        /// <param name="caller"></param>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>new instance</returns>
        /// <exception cref="ArgumentNullException">throws when comProxy is null</exception>
        /// <exception cref="DuckException">throws when its failed to compile an implementation</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception> 
        public T CreateDuckObjectFromComProxy<T>(ICOMObject caller, object comProxy) where T : ICOMObject
        {
            return (T)CreateDuckObjectFromComProxy(null, comProxy, typeof(T));
        }

        /// <summary>
        /// Creates a new duck typing instance by given generic type argument
        /// </summary>
        /// <typeparam name="T">interface result type</typeparam>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>new instance</returns>
        /// <exception cref="ArgumentNullException">throws when comProxy is null</exception>
        /// <exception cref="DuckException">throws when its failed to compile an implementation</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception> 
        public T CreateDuckObjectFromComProxy<T>(object comProxy) where T : ICOMObject
        {
            return (T)CreateDuckObjectFromComProxy(null, comProxy, typeof(T));
        }

        /// <summary>
        /// Creates a new duck typing instance
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>new instance</returns>
        /// <exception cref="ArgumentNullException">throws when comProxy is null</exception>
        /// <exception cref="DuckException">throws when its failed to compile an implementation</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception> 
        /// <exception cref="FactoryException">throws when its failed to recieve factory info</exception> 
        public ICOMObject CreateDuckObjectFromComProxy(ICOMObject caller, object comProxy)
        {
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");

            CheckInitialize();

            IFactoryInfo factoryInfo = GetDuckFactoryInfo(caller, comProxy, true);
            string className = TypeDescriptor.GetClassName(comProxy);
            string fullClassName = factoryInfo.AssemblyNamespace + ".I" + className;

            Type wrapperInterfaceType = factoryInfo.Assembly.GetType(fullClassName, true, true);
            return CreateDuckObjectFromComProxy(caller, comProxy, wrapperInterfaceType);
        }

        /// <summary>
        /// Creates a new duck typing instance which implement the given interfaces.
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="wrapperInterfaceType">interface which is implemented by the returning instance. the interface must inherit from ICOMObject</param>
        /// <returns>new instance</returns>
        /// <exception cref="ArgumentNullException">throws when comProxy, wrapperInterfaceType is null</exception>
        /// <exception cref="DuckException">throws when its failed to compile an implementation</exception>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        public ICOMObject CreateDuckObjectFromComProxy(ICOMObject caller, object comProxy, Type wrapperInterfaceType)
        {
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");
            if (null == wrapperInterfaceType)
                throw new ArgumentNullException("wrapperInterfaceType");

            Type instanceType = null;
            lock (_comObjectLock)
            {
                if (!DuckingCache.TryGetValue(wrapperInterfaceType, out instanceType))
                {
                    try
                    {
                        DuckInterface proxyInterface = new DuckInterface(wrapperInterfaceType);
                        DuckTypeGenerator implementationTypeGenerator = new DuckTypeGenerator(proxyInterface);
                        instanceType = implementationTypeGenerator.GenerateType();
                        DuckingCache.Add(wrapperInterfaceType, instanceType);
                    }
                    catch (Exception exception)
                    {
                        throw new DuckException(exception);
                    }
                }
            }

            try
            {
                ICOMObject newInstance = Activator.CreateInstance(instanceType, new object[] { this, caller, comProxy }) as ICOMObject;
                return newInstance;
            }
            catch (Exception exception)
            {
                throw new CreateInstanceException(exception);
            }        
        }
        
        /// <summary>
        /// Creates a new ICOMObject based on wrapperClassType
        /// </summary>
        /// <typeparam name="T">result type</typeparam>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="wrapperClassType">type info from wrapper class</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        public T CreateKnownObjectFromComProxy<T>(ICOMObject caller, object comProxy, Type wrapperClassType) where T:class,ICOMObject
        {
            return CreateKnownObjectFromComProxy(caller, comProxy, wrapperClassType) as T;
        }

        /// <summary>
        /// Creates a new ICOMObject based on wrapperClassType
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="wrapperClassType">type info from wrapper class</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        public ICOMObject CreateKnownObjectFromComProxy(ICOMObject caller, object comProxy, Type wrapperClassType)
        {
            if (caller.Settings.EnableKnownReferenceInspection)
            {
                return CreateObjectFromComProxy(caller, comProxy, false);
            }

            CheckInitialize();           
            try
            {
                if (null == comProxy)
                    return null;

                lock (_comObjectLock)
                {
                    // create new proxyType
                    Type comProxyType = null;
                    if (false == _proxyTypeCache.TryGetValue(wrapperClassType.FullName, out comProxyType))
                    {
                        comProxyType = comProxy.GetType();
                        _proxyTypeCache.Add(wrapperClassType.FullName, comProxyType);
                    }

                    ICOMObject newInstance = null;
                    try
                    {
                        newInstance = Activator.CreateInstance(wrapperClassType, new object[] { caller, comProxy, comProxyType }) as ICOMObject;
                        newInstance = TryReplaceInstance(caller, newInstance, comProxyType);
                    }
                    catch (Exception exception)
                    {
                        throw new CreateInstanceException(exception);
                    }

                    return newInstance;
                }               
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
        }

        /// <summary>
        /// Creates a new ICOMObject array based on wrapperClassType
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxyArray">new created proxies</param>
        /// <param name="wrapperClassType">type info from wrapper class</param>
        /// <returns>corresponding wrapper class instances or plain COMObject</returns>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        public ICOMObject[] CreateKnownObjectArrayFromComProxy(ICOMObject caller, object[] comProxyArray, Type wrapperClassType)
        {
            CheckInitialize();
            try
            {
                if (null == comProxyArray)
                    return null;

                lock (_comObjectLock)
                {
                    Type comVariantType = null;
                    ICOMObject[] newVariantArray = new ICOMObject[comProxyArray.Length];
                    for (int i = 0; i < comProxyArray.Length; i++)
                    {
                        ICOMObject newInstance = null;
                        try
                        {
                            newInstance = Activator.CreateInstance(wrapperClassType, new object[] { caller, comProxyArray[i], comVariantType }) as ICOMObject;
                            newInstance = TryReplaceInstance(caller, newInstance, comVariantType);
                        }
                        catch (Exception exception)
                        {
                            throw new CreateInstanceException(exception);
                        }
                        newVariantArray[i] = newInstance;
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
        /// Creates a new ICOMObject based on classType of comProxy. The method use Settings.EnableDynamicEventArguments to reflect dynamics
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        public ICOMObject CreateEventArgumentObjectFromComProxy(ICOMObject caller, object comProxy)
        {
            return CreateObjectFromComProxy(caller, comProxy, caller.Settings.EnableDynamicEventArguments);
        }

        /// <summary>
        /// Creates a new ICOMObject based on classType of comProxy 
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        public ICOMObject CreateObjectFromComProxy(ICOMObject caller, object comProxy, bool allowDynamicObject)
        {
            CheckInitialize();         
            try
            {
                if (null == comProxy)
                    return null;

                lock (_comObjectLock)
                {                  
                    IFactoryInfo factoryInfo = GetInstanceFactoryInfo(caller, comProxy, false);
                    if (null == factoryInfo)
                    {
                        Type comProxyType2 = null;
                        ICOMObject newInstance2 = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType2, 
                                                                            String.Empty, String.Empty, allowDynamicObject);
                        newInstance2 = TryReplaceInstance(caller, newInstance2, comProxyType2);
                        return newInstance2;
                    }
                    
                    string className = ComTypes.TypeDescriptor.GetClassName(comProxy);
                    string fullClassName = factoryInfo.AssemblyNamespace + "." + className;

                    // create new proxyType
                    Type comProxyType = null;
                    if (false == _proxyTypeCache.TryGetValue(fullClassName, out comProxyType))
                    {
                        comProxyType = comProxy.GetType();
                        _proxyTypeCache.Add(fullClassName, comProxyType);
                    }

                    ICOMObject newInstance = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType, className, fullClassName, allowDynamicObject);
                    newInstance = TryReplaceInstance(caller, newInstance, comProxyType);

                    return newInstance;
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
        /// <param name="comProxyType">Type of comProxy</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>corresponding Wrapper class Instance or plain COMObject</returns>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        public ICOMObject CreateObjectFromComProxy(ICOMObject caller, object comProxy, Type comProxyType, bool allowDynamicObject)
        {
            CheckInitialize();           
            try
            {
                if (null == comProxy)
                    return null;

                lock (_comObjectLock)
                {
                    IFactoryInfo factoryInfo = GetInstanceFactoryInfo(caller, comProxy);
                    if (null == factoryInfo)
                    {
                        Type comProxyType2 = null;
                        ICOMObject newInstance2 = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType2, String.Empty, String.Empty, allowDynamicObject);
                        newInstance2 = TryReplaceInstance(caller, newInstance2, comProxyType2);

                        return newInstance2;
                    }

                    string className = ComTypes.TypeDescriptor.GetClassName(comProxy);
                    string fullClassName = factoryInfo.AssemblyNamespace + "." + className;

                    // create new classType
                    ICOMObject newInstance = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType, className, fullClassName, allowDynamicObject);
                    newInstance = TryReplaceInstance(caller, newInstance, comProxyType);

                    return newInstance;
                }
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
        }

        /// <summary>
        /// Creates a new ICOMObject from factoryInfo
        /// </summary>
        /// <param name="factoryInfo">Factory Info from Wrapper Assemblies</param>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="comProxyType">Type of comProxy</param>
        /// <param name="className">name of COMServer proxy class</param>
        /// <param name="fullClassName">full namespace and name of COMServer proxy class</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>corresponding Wrapper class Instance or plain COMObject</returns>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        /// <exception cref="FactoryException">throws when its failed find corresponding wrapper class type</exception>
        public ICOMObject CreateObjectFromComProxy(IFactoryInfo factoryInfo, ICOMObject caller, object comProxy, 
            Type comProxyType, string className, string fullClassName, bool allowDynamicObject)
        {
            CheckInitialize();           
            try
            {
                lock (_comObjectLock)
                {
                    Type classType = null;
                    if (true == _wrapperTypeCache.TryGetValue(fullClassName, out classType))
                    {
                        // cached classType
                        ICOMObject newInstance = null;
                        try
                        {
                            newInstance = Activator.CreateInstance(classType, new object[] { caller, comProxy }) as ICOMObject;
                            newInstance = TryReplaceInstance(caller, newInstance, comProxyType);
                        }
                        catch (Exception exception)
                        {
                            throw new CreateInstanceException(exception);
                        }

                        return newInstance as ICOMObject;
                    }
                    else
                    {
                        // create new classType
                        classType = null != factoryInfo ? factoryInfo.Assembly.GetType(fullClassName, false, true) : null;
                        if (null == classType)
                            classType = RaiseResolve(caller, fullClassName, comProxyType);

                        if (null == classType)
                        {
                            if (allowDynamicObject && Settings.EnableDynamicObjects)
                            {
                                ICOMObject unkownInstance = RaiseCreateCOMDynamic(caller, comProxy);
                                if(null == unkownInstance)
                                    unkownInstance = new COMDynamicObject(caller, comProxy);
                                unkownInstance = TryReplaceInstance(caller, unkownInstance, comProxyType);
                                return unkownInstance;
                            }
                            else
                                throw new FactoryException("Class not exists: " + (true == String.IsNullOrWhiteSpace(fullClassName) ? ComTypes.TypeDescriptor.GetFullComponentClassName(comProxy) : fullClassName));
                        }

                        _wrapperTypeCache.Add(fullClassName, classType);

                        ICOMObject newInstance = null;
                        try
                        {
                            newInstance = Activator.CreateInstance(classType, new object[] { caller, comProxy, comProxyType }) as ICOMObject;
                            newInstance = TryReplaceInstance(caller, newInstance, comProxyType);
                        }
                        catch (Exception exception)
                        {
                            throw new CreateInstanceException(exception);
                        }
                        return newInstance;
                    }
                }  
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
        }

        /// <summary>
        /// Creates a new ICOMObject array
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxyArray">new created proxy array</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>corresponding Wrapper class Instance array or plain COMObject array</returns>
        /// <exception cref="CreateInstanceException">throws when its failed to create new instance</exception>
        /// <exception cref="FactoryException">throws when its failed find factory info</exception>
        public ICOMObject[] CreateObjectArrayFromComProxy(ICOMObject caller, object[] comProxyArray, bool allowDynamicObject)
        {
            CheckInitialize();
            try
            {
                if (null == comProxyArray)
                    return null;

                lock (_comObjectLock)
                {
                    Type comVariantType = null;
                    ICOMObject[] newVariantArray = new ICOMObject[comProxyArray.Length];
                    for (int i = 0; i < comProxyArray.Length; i++)
                    {
                        comVariantType = comProxyArray[i].GetType();
                        IFactoryInfo factoryInfo = GetInstanceFactoryInfo(caller, comProxyArray[i]);
                        string className = TypeDescriptor.GetClassName(comProxyArray[i]);
                        string fullClassName = factoryInfo.AssemblyNamespace + "." + className;
                        newVariantArray[i] = CreateObjectFromComProxy(factoryInfo, caller, comProxyArray[i], comVariantType, className, fullClassName, allowDynamicObject);
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
        /// Try to replace new created instance on CreateInstance event
        /// </summary>
        /// <param name="caller">parent instance</param>
        /// <param name="instance">origin instance</param>
        /// <param name="comProxyType">type of native com proxy</param>
        /// <returns>replace instance or origin instance</returns>
        private ICOMObject TryReplaceInstance(ICOMObject caller, ICOMObject instance, Type comProxyType)
        {
            Type typeToReplace = null;
            RaiseCreateInstance(instance, ref typeToReplace);
            instance.DisposeChildInstances();

            if (null != typeToReplace)
            {
                ICOMObject replaceInstance = Activator.CreateInstance(typeToReplace, new object[] { caller, instance.UnderlyingObject, comProxyType }) as ICOMObject;
                if (null != replaceInstance)
                {
                    caller.RemoveChildObject(instance);
                    RemoveObjectFromList(instance, null);
                    return replaceInstance;
                }
            }

            return instance;
        }

        #endregion

        #region Object List Methods

        /// <summary>
        /// Dispose all open objects
        /// </summary>
        public void DisposeAllCOMProxies()
        {
            lock (_globalObjectList)
            {
                // NO is appending new proxies so we free them in reverse order
                while (_globalObjectList.Count > 0)
                    _globalObjectList[_globalObjectList.Count - 1].Dispose();
                RaiseProxyCleared();
            }
        }
       
        /// <summary>
        /// Add object to global list
        /// </summary>
        /// <param name="proxy">com wrapper instance</param>
        internal void AddObjectToList(ICOMObject proxy)
        {    
            try
            {
                lock (_globalObjectList)
                {
                    _globalObjectList.Add(proxy);

                    if (HasProxyAddedRecipients)
                    {
                        IEnumerable<ICOMObject> ownerPath = GetOwnerPath(proxy);
                        RaiseProxyAdded(ownerPath, proxy);
                    }
                }               
                RaiseProxyCountChanged(_globalObjectList.Count);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
            }
        }

        /// <summary>
        /// Remove object from global list
        /// </summary>
        /// <param name="proxy">com wrapper instance</param>
        /// <param name="ownerPath">optional owner path</param>
        internal void RemoveObjectFromList(ICOMObject proxy, IEnumerable<ICOMObject> ownerPath)
        {         
            try
            {
                bool removed = false;
                lock (_globalObjectList)
                {
                    removed = _globalObjectList.Remove(proxy);

                    if (HasProxyRemovedRecipients)
                    {
                        RaiseProxyRemoved(ownerPath, proxy);
                    }
                }
                if (removed)
                    RaiseProxyCountChanged(_globalObjectList.Count);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
            }
        }

        /// <summary>
        /// Returns an array with full parent(s) path
        /// </summary>
        /// <param name="comObject">target com object</param>
        /// <returns>top down path sequence</returns>
        internal static IEnumerable<ICOMObject> GetOwnerPath(ICOMObject comObject)
        {
            if (null == comObject.ParentObject)
                return _emptyOwnerPath;

            ICOMObject parent = comObject.ParentObject;
            int parentCount = 0;
            while (null != parent)
            {
                parentCount++;
                parent = parent.ParentObject;
            }

            ICOMObject[] result = new ICOMObject[parentCount];
            parent = comObject.ParentObject;
            while (null != parent)
            {
                result[parentCount - 1] = parent;
                parentCount--;
                parent = parent.ParentObject;
            }

            return result;
        }

        /// <summary>
        /// Returns all root instances in COM proxy management
        /// </summary>
        /// <returns>Enumerable sequence of root instances</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public IEnumerable<ICOMObject> GetRootInstances()
        {
            List<ICOMObject> result = new List<ICOMObject>();
           
            try
            {
                lock (_globalObjectList)
                {
                    foreach (ICOMObject item in _globalObjectList)
                    {
                        if (null == item.ParentObject)
                            result.Add(item);
                    }
                }
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
            }

            return result;
        }

        #endregion

        #region Type Methods

        /// <summary>
        /// Analyze an object and create wrapper arround if necessary
        /// </summary>
        /// <param name="value">value as any</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>value or wrapped value</returns>
        public object WrapObject(object value, bool allowDynamicObject)
        {
            if ((null != value) && (value is MarshalByRefObject))
            {
                ICOMObject newObject = CreateObjectFromComProxy(null, value, allowDynamicObject);
                return newObject;
            }
            else
            {
                return value;
            }
        }

        /// <summary>
        /// Returns the Type for comProxy or null if param not set
        /// </summary>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>type info or null if unkown</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Type GetObjectType(object comProxy)
        {
            CheckInitialize();

            if (null == comProxy)
                return null;
            else
            {
                IFactoryInfo factoryInfo = GetInstanceFactoryInfo(null, comProxy);
                string className = TypeDescriptor.GetClassName(comProxy);
                string fullClassName = String.Format("{0}.{1}", factoryInfo.AssemblyNamespace, className);
                Type proxyType = null;
                if (!_proxyTypeCache.TryGetValue(fullClassName, out proxyType))
                {
                    proxyType = comProxy.GetType();
                    _proxyTypeCache.Add(fullClassName, proxyType);
                }
                return proxyType;
            }
        }
        
        /// <summary>
        /// Determine 2 proxies represents the same object on COM remote server
        /// </summary>
        /// <param name="obj1">object 1 to compare</param>
        /// <param name="obj2">object 2 to compare</param>
        /// <returns>true if equal, otherwise false</returns>
        public static bool EqualsOnServer(object obj1, object obj2)
        {
            return EqualsOnServer(obj1 as ICOMObject, obj2 as ICOMObject);
        }

        /// <summary>
        /// Determine 2 proxies represents the same object on COM remote server
        /// </summary>
        /// <param name="obj1">object 1 to compare</param>
        /// <param name="obj2">object 2 to compare</param>
        /// <returns>true if equal, otherwise false</returns>
        public static bool EqualsOnServer(ICOMObject obj1, ICOMObject obj2)
        {
            if (obj1.IsCurrentlyDisposing || obj1.IsDisposed)
                return ReferenceEquals(obj1, obj2);

            if (Object.ReferenceEquals(obj2, null))
                return false;

            IntPtr outValueA = IntPtr.Zero;
            IntPtr outValueB = IntPtr.Zero;
            IntPtr ptrA = IntPtr.Zero;
            IntPtr ptrB = IntPtr.Zero;
            try
            {
                ptrA = Marshal.GetIUnknownForObject(obj1.UnderlyingObject);
                int hResultA = Marshal.QueryInterface(ptrA, ref IID_IUnknown, out outValueA);

                ptrB = Marshal.GetIUnknownForObject(obj2.UnderlyingObject);
                int hResultB = Marshal.QueryInterface(ptrB, ref IID_IUnknown, out outValueB);

                return (hResultA == 0 && hResultB == 0 && ptrA == ptrB);
            }
            catch (Exception exception)
            {
                obj1.Console.WriteException(exception);
                throw exception;
            }
            finally
            {
                if (IntPtr.Zero != ptrA)
                    Marshal.Release(ptrA);

                if (IntPtr.Zero != outValueA)
                    Marshal.Release(outValueA);

                if (IntPtr.Zero != ptrB)
                    Marshal.Release(ptrB);

                if (IntPtr.Zero != outValueB)
                    Marshal.Release(outValueB);
            }
        }

        #endregion
    }
}