using System;
using System.Threading;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using COMTypes = System.Runtime.InteropServices.ComTypes;

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
    /// Creation Factory for COMObject and derived types
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Core
    {
        #region Nested

        /// <summary>
        /// Arguments in CreateInstance event
        /// </summary>
        public class OnCreateInstanceEventArgs : EventArgs
        {
            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="instance">origin instane</param>
            internal OnCreateInstanceEventArgs(ICOMObject instance)
            {
                if (null == instance)
                    throw new ArgumentNullException();
                Instance = instance;
            }

            /// <summary>
            /// DisposeChildInstance is called for the instance after event triger
            /// </summary>
            public ICOMObject Instance { get; private set; }

            /// <summary>
            /// Type muste inherit from originInstance class type and make COMObject public .ctors available
            /// </summary>
            public Type Replace { get; set; }
        }

        /// <summary>
        /// OnCreateInstance event handler
        /// </summary>
        /// <param name="sender">Core sender instance</param>
        /// <param name="args"></param>
        public delegate void OnCreateInstanceEventHandler(Core sender, OnCreateInstanceEventArgs args);

        /// <summary>
        /// ProxyCountChanged delegate
        /// </summary>
        /// <param name="proxyCount">current count of com proxies</param>
        public delegate void ProxyCountChangedHandler(int proxyCount);

        /// <summary>
        /// IsInitializedChanged delegate
        /// </summary>
        /// <param name="isInitialized"></param>
        public delegate void IsInitializedChangedHandler(bool isInitialized);

        /// <summary>
        /// Proxy added delegate
        /// </summary>
        /// <param name="sender">sender instance</param>
        /// <param name="ownerPath">comObject relation path</param>
        /// <param name="comObject">added object</param>
        public delegate void ProxyAddedHandler(Core sender, IEnumerable<ICOMObject> ownerPath, ICOMObject comObject);

        /// <summary>
        /// Proxy remove delegate
        /// </summary>
        /// <param name="sender">sender instance</param>
        /// <param name="ownerPath">former comObject relation path</param>
        /// <param name="comObject">removed object</param>
        public delegate void ProxyRemovedHandler(Core sender, IEnumerable<ICOMObject> ownerPath, ICOMObject comObject);

        /// <summary>
        /// Proxy clear delegate
        /// </summary>
        /// <param name="sender">sender instance</param>
        public delegate void ProxyClearHandler(Core sender);

        #endregion

        #region Fields

        private static Core _default;
        private bool _initalized;
        private List<ICOMObject> _globalObjectList = new List<ICOMObject>();
        private List<IFactoryInfo> _factoryList = new List<IFactoryInfo>();
        private Dictionary<string, Type> _proxyTypeCache = new Dictionary<string, Type>();
        private Dictionary<string, Type> _wrapperTypeCache = new Dictionary<string, Type>();
        private Dictionary<Guid, Guid> _hostCache = new Dictionary<Guid, Guid>();
        private Dictionary<string, Dictionary<string, string>> _entitiesListCache = new Dictionary<string, Dictionary<string, string>>();
        private List<DependentAssembly> _dependentAssemblies = new List<DependentAssembly>();
        private Assembly _thisAssembly;
        private CurrentAppDomain _appDomain;

        private object _checkInitializeLock = new object();
        private object _thisAssemblyLock = new object();
        private static object _defaultLock = new object();
        private object _factoryListLock = new object();
        private object _comObjectLock = new object();
        private static ICOMObject[] _emptyOwnerPath = new ICOMObject[0];

        private static readonly string _noAssemblyAttributeName = "NetOffice.NetOfficeAssemblyAttribute";
        private static readonly string[] _tryLoadAssemblyNames = new string[] {
                                                                                "OfficeApi.dll",
                                                                                "ExcelApi.dll",
                                                                                "WordApi.dll",
                                                                                "OutlookApi.dll",
                                                                                "PowerPointApi.dll",
                                                                                "AccessApi.dll",
                                                                                "VisioApi.dll",
                                                                                "MSProjectApi.dll",
                                                                                "PublisherApi.dll",
                                                                                "VBIDEApi.dll",
                                                                                "MSFormsApi.dll" };

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Core()
        {
            _appDomain = new CurrentAppDomain(this);
            Settings = new Settings();
            Console = new DebugConsole();
            Invoker = new Invoker(this);
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="isDefault">Mark this instance as default instance</param>
        private Core(bool isDefault)
        {
            _appDomain = new CurrentAppDomain(this);
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

        #region Properties

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
        /// The count of loaded NetOffice Assemblies
        /// </summary>
        [Category("Core"), Description("The count of loaded NetOffice Assemblies"), DisplayName("Assemblies")]
        public int AssemblyCount
        {
            get
            {
                return _factoryList.Count;
            }
        }

        /// <summary>
        /// Returns an array about currently loaded NetOffice API assemblies
        /// </summary>
        [Browsable(false)]
        public IEnumerable<IFactoryInfo> Assemblies
        {
            get
            {
                return _factoryList.ToArray();
            }
        }

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

        internal CurrentAppDomain CurrentAppDomain { get { return this._appDomain; } }

        #endregion

        #region Events

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
        /// Must be called from client assembly for COMObject Support
        /// Recieve factory infos from all loaded NetOfficeApi Assemblies in current application domain
        /// <param name="cacheOptions">NetOffice cache options</param>
        /// </summary>
        [Obsolete("Not necessary anymore(self-initializing)")]
        public void Initialize(CacheOptions cacheOptions)
        {
            Settings.CacheOptions = cacheOptions;

            if (!_initalized)
            {
                _initalized = true;
                RaiseIsInitializedChanged();
            }
                  
            bool isLocked = false;
            try
            {
                DateTime startTime = DateTime.Now;

                Monitor.Enter(_factoryListLock);
                isLocked = true;
                
                Console.WriteLine("NetOffice Core.Initialize() Core Version:{0}; Deep Loading:{1}; Load Assemblies Unsafe:{2}; AppDomain:{3}",
                     ThisAssembly.GetName().Version, Settings.EnableDeepLoading, 
                     Settings.LoadAssembliesUnsafe, AppDomain.CurrentDomain.Id.ToString() + "-" + AppDomain.CurrentDomain.FriendlyName);

                if (Settings.EnableDebugOutput)
                {
                    string localPath = UriConvert.ToLocalPath(ThisAssembly.CodeBase);
                    Console.WriteLine("Local Bind Path:{0}", localPath);
                }

                foreach (var item in _tryLoadAssemblyNames)
                    TryLoadAssembly(item);

                ClearCache();
                AddNetOfficeAssemblies();
                AddDependentNetOfficeAssemblies();

                InitializedTime = DateTime.Now - startTime;

                if (Settings.EnableDebugOutput)
                {
                    Console.WriteLine("NetOffice Core contains {0} assemblies", _factoryList.Count);
                    Console.WriteLine("NetOffice Core.Initialize() passed in {0} milliseconds", InitializedTime.TotalMilliseconds);
                }
            }
            catch (Exception throwedException)
            {
                DebugConsole.Default.WriteException(throwedException);
                throw;
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_factoryListLock);
                    isLocked = false;
                }
            }
        }

        /// <summary>
        /// Analyze assemblies in current appdomain and connect all NetOffice assemblies to the core runtime
        /// </summary>
        private void AddNetOfficeAssemblies()
        {
            var currentAssemblyName = new AssemblyName(this.ThisAssembly.FullName);
            var netOfficePublicKey = currentAssemblyName.GetPublicKeyToken();

            _dependentAssemblies.Clear();

            if (Settings.EnableDeepLoading)
            {
                Assembly[] assemblies = _appDomain.GetAssemblies();
                foreach (Assembly domainAssembly in assemblies)
                {
                    foreach (AssemblyName itemName in domainAssembly.GetReferencedAssemblies())
                    {
                        if (ContainsNetOfficePublicKeyToken(itemName, netOfficePublicKey))
                        {
                            Assembly itemAssembly = _appDomain.Load(itemName);
                            if (null == itemAssembly)
                                continue;

                            string assemblyName = itemName.Name;
                            string[] depends = AddAssembly(assemblyName, itemAssembly);
                            foreach (string depend in depends)
                            {
                                bool found = false;
                                foreach (DependentAssembly itemExistingDependency in _dependentAssemblies)
                                {
                                    if (depend == itemExistingDependency.Name)
                                    {
                                        found = true;
                                        break;
                                    }
                                }
                                if (!found)
                                    _dependentAssemblies.Add(new DependentAssembly(depend, itemAssembly));
                            }
                        }
                    }
                }
            }
            else
            {
                Assembly[] assemblies = _appDomain.GetAssemblies();
                foreach (Assembly itemAssembly in assemblies)
                {
                    string assemblyName = itemAssembly.GetName().Name;
                    if (ContainsNetOfficeAttribute(itemAssembly))
                    {
                        Console.WriteLine(string.Format("Detect NetOffice assembly {0}.", assemblyName));

                        string[] depends = AddAssembly(assemblyName, itemAssembly);
                        foreach (string depend in depends)
                        {
                            bool found = false;
                            foreach (DependentAssembly itemExistingDependency in _dependentAssemblies)
                            {
                                if (depend == itemExistingDependency.Name)
                                {
                                    found = true;
                                    break;
                                }
                            }
                            if (!found)
                                _dependentAssemblies.Add(new DependentAssembly(depend, itemAssembly));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Analyze loaded NetOffice assemblies and add dependent assemblies to the runtime if necessary
        /// </summary>
        private void AddDependentNetOfficeAssemblies()
        {
            if (!Settings.EnableAdHocLoading)
                return;

            foreach (DependentAssembly dependAssembly in _dependentAssemblies)
            {
                if (!AssemblyExistsInFactoryList(dependAssembly.Name))
                {
                    string fileName = dependAssembly.ParentAssembly.CodeBase.Substring(0, dependAssembly.ParentAssembly.CodeBase.LastIndexOf("/")) + "/" + dependAssembly.Name;
                    fileName = fileName.Replace("/", "\\").Substring(8);

                    Console.WriteLine(string.Format("Try to load dependent assembly {0}.", fileName));

                    if (System.IO.File.Exists(fileName))
                    {
                        try
                        {
                            Assembly asssembly = _appDomain.LoadFrom(fileName);
                            AddAssembly(asssembly.GetName().Name, asssembly);
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
        /// Clears proxy/type/wrapper/assembly cache etc.
        /// </summary>
        private void ClearCache()
        {
            if (CacheOptions.ClearExistingCache == Settings.CacheOptions)
            {
                _wrapperTypeCache.Clear();
                _entitiesListCache.Clear();
                _hostCache.Clear();
                _proxyTypeCache.Clear();
                _factoryList.Clear();
            }
        }

        /// <summary>
        /// Check for initialize state and call Initialize if its necessary
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public void CheckInitialize()
        {
            lock (_checkInitializeLock)
            {
                if (!_initalized)
                {
                    #pragma warning disable 612, 618
                    Initialize();
                    #pragma warning restore 612, 618
                }
            }            
        }

        /// <summary>
        /// Clears factory informations List
        /// </summary>
        public void ClearFactoryInformations()
        {
            bool isLocked = false;
            try
            {
                Monitor.Enter(_factoryListLock);
                isLocked = true;

                _factoryList.Clear();
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_factoryListLock);
                    isLocked = false;
                }
            }
        }

        /// <summary>
        /// Creates an entity support list for a proxy
        /// </summary>
        /// <param name="comProxy"></param>
        /// <returns></returns>
        internal Dictionary<string, string> GetSupportedEntities(object comProxy)
        {
            Guid parentLibraryGuid = GetParentLibraryGuid(comProxy);
            if (Guid.Empty == parentLibraryGuid)
                return null;

            string className = TypeDescriptor.GetClassName(comProxy);
            string key = (parentLibraryGuid.ToString() + className).ToLower();

            Dictionary<string, string> supportList = null;

            if (_entitiesListCache.TryGetValue(key, out supportList))
                return supportList;

            supportList = new Dictionary<string, string>();
            IDispatch dispatch = comProxy as IDispatch;
            if (null == dispatch)
                throw new COMException("Unable to cast underlying proxy to IDispatch.");

            COMTypes.ITypeInfo typeInfo = dispatch.GetTypeInfo(0, 0);
            if (null == typeInfo)
                throw new COMException("GetTypeInfo returns null.");

            IntPtr typeAttrPointer = IntPtr.Zero;
            typeInfo.GetTypeAttr(out typeAttrPointer);

            COMTypes.TYPEATTR typeAttr = (COMTypes.TYPEATTR)Marshal.PtrToStructure(typeAttrPointer, typeof(COMTypes.TYPEATTR));
            for (int i = 0; i < typeAttr.cFuncs; i++)
            {
                string strName, strDocString, strHelpFile;
                int dwHelpContext;
                IntPtr funcDescPointer = IntPtr.Zero;
                COMTypes.FUNCDESC funcDesc;
                typeInfo.GetFuncDesc(i, out funcDescPointer);
                funcDesc = (COMTypes.FUNCDESC)Marshal.PtrToStructure(funcDescPointer, typeof(COMTypes.FUNCDESC));

                switch (funcDesc.invkind)
                {
                    case COMTypes.INVOKEKIND.INVOKE_PROPERTYGET:
                    case COMTypes.INVOKEKIND.INVOKE_PROPERTYPUT:
                    case COMTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF:
                    {
                        typeInfo.GetDocumentation(funcDesc.memid, out strName, out strDocString, out dwHelpContext, out strHelpFile);
                        string outValue = "";
                        bool exists = supportList.TryGetValue("Property-" + strName, out outValue);
                        if (!exists)
                            supportList.Add("Property-" + strName, strDocString);
                        break;
                    }
                    case COMTypes.INVOKEKIND.INVOKE_FUNC:
                    {
                        typeInfo.GetDocumentation(funcDesc.memid, out strName, out strDocString, out dwHelpContext, out strHelpFile);
                        string outValue = "";
                        bool exists = supportList.TryGetValue("Method-" + strName, out outValue);
                        if (!exists)
                            supportList.Add("Method-" + strName, strDocString);
                        break;   
                    }
                }

                typeInfo.ReleaseFuncDesc(funcDescPointer);
            }

            typeInfo.ReleaseTypeAttr(typeAttrPointer);
            Marshal.ReleaseComObject(typeInfo);

            _entitiesListCache.Add(key, supportList);

            return supportList;
        }

        #endregion

        #region Create COMObject Methods

        /// <summary>
        /// Creates a new COMObject based on wrapperClassType
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="wrapperClassType">type info from wrapper class</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        public ICOMObject CreateKnownObjectFromComProxy(ICOMObject caller, object comProxy, Type wrapperClassType)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                if (null == comProxy)
                    return null;

                Monitor.Enter(_comObjectLock);
                isLocked = true;

                // create new proxyType
                Type comProxyType = null;
                if (false == _proxyTypeCache.TryGetValue(wrapperClassType.FullName, out comProxyType))
                {
                    comProxyType = comProxy.GetType();
                    _proxyTypeCache.Add(wrapperClassType.FullName, comProxyType);
                }

                ICOMObject newInstance = Activator.CreateInstance(wrapperClassType, new object[] { caller, comProxy, comProxyType }) as ICOMObject;
                newInstance = TryReplaceInstance(caller, newInstance, comProxyType);

                return newInstance;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_comObjectLock);
                    isLocked = false;
                }
            }
        }

        /// <summary>
        /// Creates a new COMObject array based on wrapperClassType
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxyArray">new created proxies</param>
        /// <param name="wrapperClassType">type info from wrapper class</param>
        /// <returns>corresponding wrapper class instances or plain COMObject</returns>
        public ICOMObject[] CreateKnownObjectArrayFromComProxy(ICOMObject caller, object[] comProxyArray, Type wrapperClassType)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                if (null == comProxyArray)
                    return null;

                Monitor.Enter(_comObjectLock);
                isLocked = true;

                Type comVariantType = null;
                ICOMObject[] newVariantArray = new ICOMObject[comProxyArray.Length];
                for (int i = 0; i < comProxyArray.Length; i++)
                {
                    ICOMObject newInstance = Activator.CreateInstance(wrapperClassType, new object[] { caller, comProxyArray[i], comVariantType }) as ICOMObject;
                    newInstance = TryReplaceInstance(caller, newInstance, comVariantType);
                    newVariantArray[i] = newInstance;
                }

                return newVariantArray;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_comObjectLock);
                    isLocked = false;
                }
            }
        }

        /// <summary>
        /// Creates a new COMObject based on classType of comProxy 
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        public ICOMObject CreateObjectFromComProxy(ICOMObject caller, object comProxy)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                if (null == comProxy)
                    return null;

                Monitor.Enter(_comObjectLock);
                isLocked = true;

                IFactoryInfo factoryInfo = GetFactoryInfo(comProxy);
                if (null == factoryInfo)
                {
                    Type comProxyType2 = null;
                    ICOMObject newInstance2 = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType2, "", "");
                    newInstance2 = TryReplaceInstance(caller, newInstance2, comProxyType2);

                    return newInstance2;
                }

                string className = TypeDescriptor.GetClassName(comProxy);
                string fullClassName = factoryInfo.AssemblyNamespace + "." + className;

                // create new proxyType
                Type comProxyType = null;
                if (false == _proxyTypeCache.TryGetValue(fullClassName, out comProxyType))
                {
                    comProxyType = comProxy.GetType();
                    _proxyTypeCache.Add(fullClassName, comProxyType);
                }

                ICOMObject newInstance = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType, className, fullClassName);
                newInstance = TryReplaceInstance(caller, newInstance, comProxyType);

                return newInstance;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_comObjectLock);
                    isLocked = false;
                }
            }
        }

        /// <summary>
        /// Creates a new COMObject based on classType of comProxy 
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="comProxyType">Type of comProxy</param>
        /// <returns>corresponding Wrapper class Instance or plain COMObject</returns>
        public ICOMObject CreateObjectFromComProxy(ICOMObject caller, object comProxy, Type comProxyType)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                if (null == comProxy)
                    return null;

                Monitor.Enter(_comObjectLock);
                isLocked = true;

                IFactoryInfo factoryInfo = GetFactoryInfo(comProxy);
                if (null == factoryInfo)
                {
                    Type comProxyType2 = null;
                    ICOMObject newInstance2 = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType2, "", "");
                    newInstance2 = TryReplaceInstance(caller, newInstance2, comProxyType2);

                    return newInstance2;
                }

                string className = TypeDescriptor.GetClassName(comProxy);
                string fullClassName = factoryInfo.AssemblyNamespace + "." + className;

                // create new classType
                ICOMObject newInstance = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType, className, fullClassName);
                newInstance = TryReplaceInstance(caller, newInstance, comProxyType);

                return newInstance;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_comObjectLock);
                    isLocked = false;
                }
            }
        }

        /// <summary>
        /// Creates a new COMObject from factoryInfo
        /// </summary>
        /// <param name="factoryInfo">Factory Info from Wrapper Assemblies</param>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="comProxyType">Type of comProxy</param>
        /// <param name="className">name of COMServer proxy class</param>
        /// <param name="fullClassName">full namespace and name of COMServer proxy class</param>
        /// <returns>corresponding Wrapper class Instance or plain COMObject</returns>
        public ICOMObject CreateObjectFromComProxy(IFactoryInfo factoryInfo, ICOMObject caller, object comProxy, Type comProxyType, string className, string fullClassName)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                Monitor.Enter(_comObjectLock);
                isLocked = true;

                Type classType = null;
                if (true == _wrapperTypeCache.TryGetValue(fullClassName, out classType))
                {
                    // cached classType
                    ICOMObject newInstance = Activator.CreateInstance(classType, new object[] { caller, comProxy }) as ICOMObject;
                    newInstance = TryReplaceInstance(caller, newInstance, comProxyType);
                    return newInstance as ICOMObject;
                }
                else
                {
                    // create new classType
                    classType = null != factoryInfo ? factoryInfo.Assembly.GetType(fullClassName, false, true) : null;
                    if (null == classType)
                    {
                        if (Settings.EnableDynamicObjects)
                        {
                            ICOMObject unkownInstance = new COMDynamicObject(caller, comProxy);
                            unkownInstance = TryReplaceInstance(caller, unkownInstance, comProxyType);
                            return unkownInstance;
                        }
                        else
                            throw new ArgumentException("Class not exists: " + fullClassName);
                    }

                    _wrapperTypeCache.Add(fullClassName, classType);
                    ICOMObject newInstance = Activator.CreateInstance(classType, new object[] { caller, comProxy, comProxyType }) as ICOMObject;
                    newInstance = TryReplaceInstance(caller, newInstance, comProxyType);
                    return newInstance;
                }
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_comObjectLock);
                    isLocked = false;
                }
            }
        }

        /// <summary>
        /// Creates a new COMObject array
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxyArray">new created proxy array</param>
        /// <returns>corresponding Wrapper class Instance array or plain COMObject array</returns>
        public ICOMObject[] CreateObjectArrayFromComProxy(ICOMObject caller, object[] comProxyArray)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                if (null == comProxyArray)
                    return null;

                Monitor.Enter(_comObjectLock);
                isLocked = true;

                Type comVariantType = null;
                ICOMObject[] newVariantArray = new ICOMObject[comProxyArray.Length];
                for (int i = 0; i < comProxyArray.Length; i++)
                {
                    comVariantType = comProxyArray[i].GetType();
                    IFactoryInfo factoryInfo = GetFactoryInfo(comProxyArray[i]);
                    string className = TypeDescriptor.GetClassName(comProxyArray[i]);
                    string fullClassName = factoryInfo.AssemblyNamespace + "." + className;
                    newVariantArray[i] = CreateObjectFromComProxy(factoryInfo, caller, comProxyArray[i], comVariantType, className, fullClassName);
                }
                return newVariantArray;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw;
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_comObjectLock);
                    isLocked = false;
                }
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
            bool isLocked = false;
            try
            {
                Monitor.Enter(_globalObjectList);
                isLocked = true;

                _globalObjectList.Add(proxy);

                if (HasProxyAddedRecipients)
                {
                    IEnumerable<ICOMObject> ownerPath = GetOwnerPath(proxy);
                    RaiseProxyAdded(ownerPath, proxy);
                }

                RaiseProxyCountChanged(_globalObjectList.Count);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_globalObjectList);
                    isLocked = false;
                }
            }
        }

        /// <summary>
        /// Remove object from global list
        /// </summary>
        /// <param name="proxy">com wrapper instance</param>
        /// <param name="ownerPath">optional owner path</param>
        internal void RemoveObjectFromList(ICOMObject proxy, IEnumerable<ICOMObject> ownerPath)
        {
            bool isLocked = false;
            try
            {
                Monitor.Enter(_globalObjectList);
                isLocked = true;

                _globalObjectList.Remove(proxy);

                if (HasProxyRemovedRecipients)
                {                 
                    RaiseProxyRemoved(ownerPath, proxy);
                }

                RaiseProxyCountChanged(_globalObjectList.Count);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_globalObjectList);
                    isLocked = false;
                }
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

        #endregion

        #region Private Methods

        /// <summary>
        /// Returns info the assembly is a NetOffice Api Assembly
        /// </summary>
        /// <param name="itemAssembly">assembly informations</param>
        /// <returns>true if NetOffice assembly, otherwise false</returns>
        private bool ContainsNetOfficeAttribute(Assembly itemAssembly)
        {
            try
            {
                List<string> dependAssemblies = new List<string>();
                object[] attributes = itemAssembly.GetCustomAttributes(true);
                foreach (object itemAttribute in attributes)
                {
                    string fullnameAttribute = itemAttribute.GetType().FullName;
                    if (fullnameAttribute == _noAssemblyAttributeName)
                        return true;
                }
                return false;
            }
            catch (System.IO.FileNotFoundException exception)
            {
                Console.WriteException(exception);
                return false;
            }
        }

        /// <summary>
        /// Returns info the assembly is a NetOffice Api Assembly with known keytoken
        /// </summary>
        /// <param name="itemName">Assembly name to check for equal public key in strong name.</param>
        /// <param name="netOfficePublicKeyToken">NetOffice public key token.</param>
        /// <returns>true if NetOffice assembly with token, otherwise false</returns>
        internal bool ContainsNetOfficePublicKeyToken(AssemblyName itemName, byte[] netOfficePublicKeyToken)
        {
            var assemblyToken = itemName.GetPublicKeyToken();
            if (assemblyToken == null)
            {
                return false;
            }

            return netOfficePublicKeyToken.SequenceEqual(assemblyToken);
        }

        /// <summary>
        /// Check for loaded assembly in factory list
        /// </summary>
        /// <param name="name">name of the assembly</param>
        /// <returns>true if exists, otherwise false</returns>
        private bool AssemblyExistsInFactoryList(string name)
        {
            if (name.EndsWith(".dll", StringComparison.InvariantCultureIgnoreCase))
                name = name.Substring(0, name.Length - 4);

            foreach (IFactoryInfo item in _factoryList)
            {
                if (item.Assembly.GetName().Name.StartsWith(name, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Add assembly to list
        /// </summary>
        /// <param name="name">name of the assembly</param>
        /// <param name="itemAssembly">assemmbly to add</param>
        /// <returns>list of dependend assemblies</returns>
        private string[] AddAssembly(string name, Assembly itemAssembly)
        {
            List<string> dependAssemblies = new List<string>();
            object[] attributes = itemAssembly.GetCustomAttributes(true);
            foreach (object itemAttribute in attributes)
            {
                string fullnameAttribute = itemAttribute.GetType().FullName;
                if (fullnameAttribute == "NetOffice.NetOfficeAssemblyAttribute")
                {
                    bool exists = false;
                    NetOffice.IFactoryInfo factoryInfo = null;

                    foreach (IFactoryInfo itemFactory in _factoryList)
                    {
                        if (("NetOffice." + name).Equals(itemFactory.AssemblyNamespace, StringComparison.InvariantCultureIgnoreCase))
                        {
                            factoryInfo = itemFactory;
                            exists = true;
                            break;
                        }
                    }

                    if (null == factoryInfo)
                    {
                        Type factoryInfoType = itemAssembly.GetType(name + ".Utils.ProjectInfo");
                        object utilsResult = Activator.CreateInstance(factoryInfoType);
                        if (null == utilsResult)
                            throw new NetOfficeException(String.Format("Unable to create {0} factory info", name));
                        factoryInfo = utilsResult as IFactoryInfo;
                        if (null == factoryInfo)
                        {
                            throw new NetOfficeException(String.Format("Unexpected {0} factory info. Assembly {0}", name, itemAssembly));
                        }
                        foreach (IFactoryInfo itemFactory in _factoryList)
                        {
                            if (itemFactory.Assembly.FullName == factoryInfo.Assembly.FullName)
                            {
                                exists = true;
                                break;
                            }
                        }
                    }

                    if (null == factoryInfo)
                        throw new NetOfficeException(String.Format("Unable to find {0} factory info", name));

                    if (!exists)
                    {
                        _factoryList.Add(factoryInfo);
                        Console.WriteLine("NetOffice Core recieved IFactoryInfo:{0}:{1}", factoryInfo.Assembly.FullName, factoryInfo.Assembly.FullName);
                    }

                    foreach (string itemDependency in factoryInfo.Dependencies)
                    {
                        bool found = false;
                        foreach (string itemExistingDependency in dependAssemblies)
                        {
                            if (itemDependency == itemExistingDependency)
                            {
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                            dependAssemblies.Add(itemDependency);
                    }
                }
            }

            return dependAssemblies.ToArray();
        }

        /// <summary>
        /// Returns id of an interface
        /// </summary>
        /// <param name="typeInfo">com type informations</param>
        /// <returns>internface id(iid)</returns>
        internal static Guid GetTypeGuid(COMTypes.ITypeInfo typeInfo)
        {
            IntPtr attribPtr = IntPtr.Zero;
            typeInfo.GetTypeAttr(out attribPtr);
            COMTypes.TYPEATTR Attributes = (COMTypes.TYPEATTR)Marshal.PtrToStructure(attribPtr, typeof(COMTypes.TYPEATTR));
            Guid typeGuid = Attributes.guid;
            typeInfo.ReleaseTypeAttr(attribPtr);
            return typeGuid;
        }
          
        /// <summary>
        /// Get the guid from type lib there is the type defined
        /// </summary>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>guid from containing component</returns>
        public Guid GetParentLibraryGuid(object comProxy)
        {
            if (null == comProxy)
                throw new ArgumentNullException();

            IDispatch dispatcher = comProxy as IDispatch;
            if (null == dispatcher)
                return Guid.Empty;

            COMTypes.ITypeInfo typeInfo = dispatcher.GetTypeInfo(0, 0);
            COMTypes.ITypeLib parentTypeLib = null;

            Guid typeGuid = GetTypeGuid(typeInfo);
            Guid parentGuid = Guid.Empty;

            if (!_hostCache.TryGetValue(typeGuid, out parentGuid))
            {
                int i = 0;
                typeInfo.GetContainingTypeLib(out parentTypeLib, out i);

                IntPtr attributesPointer = IntPtr.Zero;
                parentTypeLib.GetLibAttr(out attributesPointer);

                COMTypes.TYPELIBATTR attributes = (COMTypes.TYPELIBATTR)Marshal.PtrToStructure(attributesPointer, typeof(COMTypes.TYPELIBATTR));
                parentGuid = attributes.guid;
                parentTypeLib.ReleaseTLibAttr(attributesPointer);
                Marshal.ReleaseComObject(parentTypeLib);

                _hostCache.Add(typeGuid, parentGuid);
            }

            Marshal.ReleaseComObject(typeInfo);

            return parentGuid;
        }

        /// <summary>
        /// Get wrapper class factory info 
        /// </summary>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>factory info from corresponding assembly</returns>
        private IFactoryInfo GetFactoryInfo(object comProxy)
        {
            if (_factoryList.Count == 0)
                return null;

            string className = TypeDescriptor.GetClassName(comProxy);
            Guid hostGuid = GetParentLibraryGuid(comProxy);

            foreach (IFactoryInfo item in _factoryList)
            {
                foreach (var guid in item.ComponentGuid)
                    if (true == guid.Equals(hostGuid))
                        return item;
            }

            // failback because some types was multiple defined (not allowed in COM but in fact MS do this)
            foreach (IFactoryInfo item in _factoryList)
            {
                if (item.Contains(className))
                    return item;
            }

            string message = string.Format("Class {0}:{1} not found in loaded NetOffice Assemblies{2}", hostGuid, className, Environment.NewLine);
            message += string.Format("Currently loaded NetOfficeApi Assemblies{0}", Environment.NewLine);
            foreach (IFactoryInfo item in _factoryList)
                message += string.Format("Loaded NetOffice Assembly:{0} {1}{2}", item.ComponentGuid, item.Assembly.FullName, Environment.NewLine);

            throw new NetOfficeException(message);
        }

        /// <summary>
        /// Assembly loader for (multitargeting) host scenarios
        /// </summary>
        /// <param name="fileName">full file name</param>
        /// <returns>assembly instance or null</returns>
        private Assembly TryLoadAssembly(string fileName)
        {
            try
            {
                string localAssemblyPath = UriConvert.ToLocalPath(ThisAssembly.CodeBase);
                string directoryName = System.IO.Path.GetDirectoryName(localAssemblyPath); 
                string fullFileName = System.IO.Path.Combine(directoryName, fileName);

                if (System.IO.File.Exists(fullFileName))
                {
                    Assembly assembly = _appDomain.LoadFrom(fullFileName);
                    if (null != assembly)
                    { 
                        Type factoryInfoType = assembly.GetType(fileName.Substring(0, fileName.Length - 4) + ".Utils.ProjectInfo", false, false);
                        NetOffice.IFactoryInfo factoryInfo = Activator.CreateInstance(factoryInfoType) as NetOffice.IFactoryInfo;
                        bool exists = false;
                        foreach (IFactoryInfo itemFactory in _factoryList)
                        {
                            if (itemFactory.Assembly.FullName == factoryInfo.Assembly.FullName)
                            {
                                exists = true;
                                break;
                            }
                        }
                        if (!exists)
                        {
                            _factoryList.Add(factoryInfo);
                            if(Settings.EnableDebugOutput)
                                Console.WriteLine("NetOffice Core Recieved IFactoryInfo:{0}:{1}", factoryInfo.Assembly.FullName, factoryInfo.Assembly.FullName);
                        }
                    }
                    return assembly;
                }
                else
                {
                    if (Settings.EnableDebugOutput)
                        Console.WriteLine("Unable To Find {0}", fullFileName);
                    return null;
                }
            }
            catch (Exception exception)
            {
                Console.WriteException(exception);
                return null;
            }
        }

        #endregion

        #region Type

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
                IFactoryInfo factoryInfo = GetFactoryInfo(comProxy);
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

        #endregion
    }
}