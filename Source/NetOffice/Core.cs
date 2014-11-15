using System;
using System.Threading;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using COMTypes = System.Runtime.InteropServices.ComTypes;

namespace NetOffice
{
    #region IDispatch

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
    public class Core
    {
        #region Fields

        private bool _initalized;
        private bool _assemblyResolveEventConnected;
        private List<COMObject> _globalObjectList = new List<COMObject>();
        private List<IFactoryInfo> _factoryList = new List<IFactoryInfo>();
        private Dictionary<string, Type> _proxyTypeCache = new Dictionary<string, Type>();
        private Dictionary<string, Type> _wrapperTypeCache = new Dictionary<string, Type>();
        private Dictionary<Guid, Guid> _hostCache = new Dictionary<Guid, Guid>();
        private Dictionary<string, Dictionary<string, string>> _entitiesListCache = new Dictionary<string, Dictionary<string, string>>();
        private List<DependentAssembly> _dependentAssemblies = new List<DependentAssembly>();

        private static object _factoryListLock = new object();
        private static object _comObjectLock = new object();
        private static object _globalObjectListLock = new object();

        private static Assembly _thisAssembly = Assembly.GetAssembly(typeof(COMObject));
        private static string[] _knownNetOfficeKeyTokens;
        private static object _defaultLock = new object();

        private static readonly string _noAssemblyAttributeName = "NetOffice.NetOfficeAssemblyAttribute";
        private static readonly string[] _tryLoadAssemblyNames = new string[] { "ExcelApi.dll", "WordApi.dll", "OutlookApi.dll", "PowerPointApi.dll", "AccessApi.dll", "VisioApi.dll", "MSProjectApi.dll" };

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Core()
        {
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
        /// Returns info about intialized state
        /// </summary>
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
        private static Core _default;

        /// <summary>
        /// Core Settings
        /// </summary>
        public Settings Settings { get; internal set; }

        /// <summary>
        /// Core Console
        /// </summary>
        public DebugConsole Console { get; internal set; }

        /// <summary>
        /// Core Invoker
        /// </summary>
        public Invoker Invoker { get; internal set; }

        /// <summary>
        /// Returns the instance ist the shared default core
        /// </summary>
        public bool IsDefault { get; private set; }

        /// <summary>
        /// returns an array about currently loaded NetOfficeApi assemblies
        /// </summary>
        public IFactoryInfo[] Assemblies
        {
            get
            {
                return _factoryList.ToArray();
            }
        }

        /// <summary>
        /// Returns current count of open proxies
        /// </summary>
        public int ProxyCount
        {
            get
            {
                return _globalObjectList.Count;
            }
        }

        #endregion

        #region Events

        /// <summary>
        /// ProxyCountChanged delegate
        /// </summary>
        /// <param name="proxyCount">current count of com proxies</param>
        public delegate void ProxyCountChangedHandler(int proxyCount);

        /// <summary>
        /// notify info the count of proxies there open are changed
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

            try
            {
                if (Settings.EnableProxyCountChannel)
                    this.Console.SendPipeChannelMessage(Settings.ProxyCountChannelName, String.Format("Current Proxy Count:{0}", proxyCount));
            }
            catch (Exception throwedException)
            {
                DebugConsole.Default.WriteException(throwedException);
                throw;
            }
        }

        #endregion

        #region Factory Methods

        /// <summary>
        /// Must be called from client assembly for COMObject Support
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
            _initalized = true;
            bool isLocked = false;
            try
            {
                Monitor.Enter(_factoryListLock);
                isLocked = true;

                Console.WriteLine("NetOffice Core.Initialize() NO Version:{1} DeepLevel:{0}", Settings.EnableDeepLoading, this.GetType().Assembly.GetName().Version);

                foreach (var item in _tryLoadAssemblyNames)
                    TryLoadAssembly(item);

                if (!_assemblyResolveEventConnected)
                {
                    AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
                    _assemblyResolveEventConnected = true;
                }

                ClearCache();
                AddNetOfficeAssemblies();
                AddDependentNetOfficeAssemblies();

                Console.WriteLine("Factory contains {0} assemblies", _factoryList.Count);
                Console.WriteLine("NetOffice Core.Initialize() passed");
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
        /// analyze assemblies in current appdomain and connect all NetOffice assemblies to the core runtime
        /// </summary>
        private void AddNetOfficeAssemblies()
        {
            _dependentAssemblies.Clear();

            if (Settings.EnableDeepLoading)
            {
                Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();
                foreach (Assembly domainAssembly in assemblies)
                {
                    foreach (AssemblyName itemName in domainAssembly.GetReferencedAssemblies())
                    {
                        if (ContainsNetOfficePublicKeyToken(itemName))
                        {
                            string assemblyName = itemName.Name;
                            Assembly itemAssembly = Assembly.Load(itemName);

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
                Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();
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
        /// analyze loaded NetOffice assemblies and add dependent assemblies to the runtime if necessary
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
                            Assembly asssembly = Assembly.LoadFile(fileName);
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
        /// clears proxy/type/wrapper/assembly cache etc.
        /// </summary>
        private void ClearCache()
        {
            // clear entities cache
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
        /// Check for inialize state and call Initialze if its necessary
        /// </summary>
        internal void CheckInitialize()
        {
            if (!_initalized)
                Initialize();
        }

        /// <summary>
        /// clears factory informations List
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
        /// creates an entity support list for a proxy
        /// </summary>
        /// <param name="comProxy"></param>
        /// <returns></returns>
        internal Dictionary<string, string> GetSupportedEntities(object comProxy)
        {
            Guid parentLibraryGuid = GetParentLibraryGuid(comProxy);
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
                System.Runtime.InteropServices.ComTypes.FUNCDESC funcDesc;
                typeInfo.GetFuncDesc(i, out funcDescPointer);
                funcDesc = (COMTypes.FUNCDESC)Marshal.PtrToStructure(funcDescPointer, typeof(System.Runtime.InteropServices.ComTypes.FUNCDESC));

                switch (funcDesc.invkind)
                {
                    case System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYGET:
                    case System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYPUT:
                    case System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF:
                    {
                        typeInfo.GetDocumentation(funcDesc.memid, out strName, out strDocString, out dwHelpContext, out strHelpFile);
                        string outValue = "";
                        bool exists = supportList.TryGetValue("Property-" + strName, out outValue);
                        if (!exists)
                            supportList.Add("Property-" + strName, strDocString);
                        break;
                    }
                    case System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_FUNC:
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
        /// creates a new COMObject based on wrapperClassType
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="wrapperClassType">type info from wrapper class</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        public COMObject CreateKnownObjectFromComProxy(COMObject caller, object comProxy, Type wrapperClassType)
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

                COMObject newClass = Activator.CreateInstance(wrapperClassType, new object[] { caller, comProxy, comProxyType }) as COMObject;
                return newClass;
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
        /// creates a new COMObject array based on wrapperClassType
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxyArray">new created proxies</param>
        /// <param name="wrapperClassType">type info from wrapper class</param>
        /// <returns>corresponding wrapper class instances or plain COMObject</returns>
        public COMObject[] CreateKnownObjectArrayFromComProxy(COMObject caller, object[] comProxyArray, Type wrapperClassType)
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
                COMObject[] newVariantArray = new COMObject[comProxyArray.Length];
                for (int i = 0; i < comProxyArray.Length; i++)
                    newVariantArray[i] = Activator.CreateInstance(wrapperClassType, new object[] { caller, comProxyArray[i], comVariantType }) as COMObject;

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
        /// creates a new COMObject based on classType of comProxy 
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>corresponding wrapper class instance or plain COMObject</returns>
        public COMObject CreateObjectFromComProxy(COMObject caller, object comProxy)
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
                string className = TypeDescriptor.GetClassName(comProxy);
                string fullClassName = factoryInfo.AssemblyNamespace + "." + className;

                // create new proxyType
                Type comProxyType = null;
                if (false == _proxyTypeCache.TryGetValue(fullClassName, out comProxyType))
                {
                    comProxyType = comProxy.GetType();
                    _proxyTypeCache.Add(fullClassName, comProxyType);
                }

                COMObject newObject = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType, className, fullClassName);
                return newObject;
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
        /// creates a new COMObject based on classType of comProxy 
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="comProxyType">Type of comProxy</param>
        /// <returns>corresponding Wrapper class Instance or plain COMObject</returns>
        public COMObject CreateObjectFromComProxy(COMObject caller, object comProxy, Type comProxyType)
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

                string className = TypeDescriptor.GetClassName(comProxy);
                string fullClassName = factoryInfo.AssemblyNamespace + "." + className;

                // create new classType
                COMObject newObject = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType, className, fullClassName);
                return newObject;
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
        /// creates a new COMObject from factoryInfo
        /// </summary>
        /// <param name="factoryInfo">Factory Info from Wrapper Assemblies</param>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="comProxyType">Type of comProxy</param>
        /// <param name="className">name of COMServer proxy class</param>
        /// <param name="fullClassName">full namespace and name of COMServer proxy class</param>
        /// <returns>corresponding Wrapper class Instance or plain COMObject</returns>
        public COMObject CreateObjectFromComProxy(IFactoryInfo factoryInfo, COMObject caller, object comProxy, Type comProxyType, string className, string fullClassName)
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
                    object newClass = Activator.CreateInstance(classType, new object[] { caller, comProxy });
                    return newClass as COMObject;
                }
                else
                {
                    // create new classType
                    classType = factoryInfo.Assembly.GetType(fullClassName, false, true);
                    if (null == classType)
                        throw new ArgumentException("Class not exists: " + fullClassName);

                    _wrapperTypeCache.Add(fullClassName, classType);
                    COMObject newClass = Activator.CreateInstance(classType, new object[] { caller, comProxy, comProxyType }) as COMObject;
                    return newClass;
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
        ///  creates a new COMObject array
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxyArray">new created proxy array</param>
        /// <returns>corresponding Wrapper class Instance array or plain COMObject array</returns>
        public COMObject[] CreateObjectArrayFromComProxy(COMObject caller, object[] comProxyArray)
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
                COMObject[] newVariantArray = new COMObject[comProxyArray.Length];
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

        #endregion

        #region Object List Methods

        /// <summary>
        /// dispose all open objects
        /// </summary>
        public void DisposeAllCOMProxies()
        {
            // NO is appending new proxies so we free them in reverse order
            while (_globalObjectList.Count > 0)
                _globalObjectList[_globalObjectList.Count -1].Dispose();
        }

        /// <summary>
        /// add object to global list
        /// </summary>
        /// <param name="proxy">com wrapper instance</param>
        internal void AddObjectToList(COMObject proxy)
        {
            bool isLocked = false;
            try
            {
                Monitor.Enter(_globalObjectList);
                isLocked = true;

                _globalObjectList.Add(proxy);

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
        /// remove object from global list
        /// </summary>
        /// <param name="proxy">com wrapper instance</param>
        internal void RemoveObjectFromList(COMObject proxy)
        {
            bool isLocked = false;
            try
            {
                Monitor.Enter(_globalObjectList);
                isLocked = true;

                _globalObjectList.Remove(proxy);

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

        #endregion

        #region Private Methods

        /// <summary>
        /// returns info the assembly is a NetOffice Api Assembly
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
        /// returns info the assembly is a NetOffice Api Assembly with known keytoken
        /// </summary>
        /// <param name="itemName">assembly informations</param>
        /// <returns>true if NetOffice assembly with token, otherwise false</returns>
        private bool ContainsNetOfficePublicKeyToken(AssemblyName itemName)
        {
            try
            {
                string targetKeyToken = itemName.FullName.Substring(itemName.FullName.LastIndexOf(" ") + 1);
                foreach (string item in KnownNetOfficeKeyTokens)
                {
                    if (item.EndsWith(targetKeyToken, StringComparison.InvariantCultureIgnoreCase))
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
        /// contains a list of all known netoffice 
        /// </summary>
        private static string[] KnownNetOfficeKeyTokens
        {
            get
            {
                if (null == _knownNetOfficeKeyTokens)
                {
                    Type thisType = typeof(Core);
                    System.IO.Stream ressourceStream = thisType.Assembly.GetManifestResourceStream(thisType.Namespace + ".KeyTokens.txt");
                    System.IO.StreamReader textStreamReader = new System.IO.StreamReader(ressourceStream);
                    string text = textStreamReader.ReadToEnd();
                    ressourceStream.Close();
                    textStreamReader.Close();
                    _knownNetOfficeKeyTokens = text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                }
                return _knownNetOfficeKeyTokens;
            }
        }

        /// <summary>
        /// check for loaded assembly in factory list
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
        /// add assembly to list
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
                        Console.WriteLine("Recieve IFactoryInfo:{0}:{1}", factoryInfo.Assembly.FullName, factoryInfo.Assembly.FullName);
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
        /// returns id of an interface
        /// </summary>
        /// <param name="typeInfo">com type informations</param>
        /// <returns>internface id(iid)</returns>
        private static Guid GetTypeGuid(COMTypes.ITypeInfo typeInfo)
        {
            IntPtr attribPtr = IntPtr.Zero;
            typeInfo.GetTypeAttr(out attribPtr);
            COMTypes.TYPEATTR Attributes = (COMTypes.TYPEATTR)Marshal.PtrToStructure(attribPtr, typeof(COMTypes.TYPEATTR));
            Guid typeGuid = Attributes.guid;
            typeInfo.ReleaseTypeAttr(attribPtr);
            return typeGuid;
        }

        /// <summary>
        /// get the guid from type lib there is the type defined
        /// </summary>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>guid from containing component</returns>
        public Guid GetParentLibraryGuid(object comProxy)
        {
            IDispatch dispatcher = comProxy as IDispatch;
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
        /// get wrapper class factory info 
        /// </summary>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>factory info from corresponding assembly</returns>
        private IFactoryInfo GetFactoryInfo(object comProxy)
        {
            if (_factoryList.Count == 0)
            {
                string notInitMessage = "Factory is not initialized with NetOffice assemblies.";
                throw new NetOfficeException(notInitMessage);
            }

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
        /// Occurs when the resolution of an assembly fails.
        /// </summary>
        /// <param name="sender">The source of the event</param>
        /// <param name="args">A System.ResolveEventArgs that contains the event data</param>
        /// <returns>The System.Reflection.Assembly that resolves the type, assembly, or resource or null if the assembly cannot be resolved</returns>
        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                // dont care for resources
                if ((!String.IsNullOrEmpty(args.Name) && args.Name.ToLower().Trim().IndexOf(".resources") > -1))
                    return null;

                string directoryName = _thisAssembly.CodeBase.Substring(0, _thisAssembly.CodeBase.LastIndexOf("/"));
                directoryName = directoryName.Replace("/", "\\").Substring(8);
                string fileName = args.Name.Substring(0, args.Name.IndexOf(","));
                string fullFileName = System.IO.Path.Combine(directoryName, fileName + ".dll");
                if (System.IO.File.Exists(fullFileName))
                {
                    Console.WriteLine(string.Format("Try to resolve assembly {0}", args.Name));
                    Assembly assembly = System.Reflection.Assembly.Load(args.Name);
                    return assembly;
                }
                else
                {
                    Console.WriteLine(string.Format("Unable to resolve assembly {0}. The file doesnt exists in current codebase.", args.Name));
                    return null;
                }
            }
            catch (Exception exception)
            {
                Console.WriteException(exception);
                return null;
            }
        }

        /// <summary>
        /// Assembly loader for multitargeting(host) scenarios
        /// </summary>
        /// <param name="fileName">full file name</param>
        /// <returns>assembly instance or null</returns>
        private Assembly TryLoadAssembly(string fileName)
        {
            try
            {
                string directoryName = _thisAssembly.CodeBase.Substring(0, _thisAssembly.CodeBase.LastIndexOf("/"));
                directoryName = directoryName.Replace("/", "\\").Substring(8);
                string fullFileName = System.IO.Path.Combine(directoryName, fileName);
                if (System.IO.File.Exists(fullFileName))
                {

                    Assembly assembly = System.Reflection.Assembly.LoadFrom(fullFileName);
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
                        Console.WriteLine("Recieve IFactoryInfo:{0}:{1}", factoryInfo.Assembly.FullName, factoryInfo.Assembly.FullName);
                    }
                    return assembly;
                }
                else
                {
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
        /// returns the Type for comProxy or null if param not set
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