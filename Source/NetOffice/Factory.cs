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

    [Guid("00020400-0000-0000-c000-000000000046"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDispatch
    {
        [PreserveSig]
        int GetTypeInfoCount();

        System.Runtime.InteropServices.ComTypes.ITypeInfo GetTypeInfo([MarshalAs(UnmanagedType.U4)] int iTInfo, [MarshalAs(UnmanagedType.U4)] int lcid);

        [PreserveSig]
        int GetIDsOfNames(ref Guid riid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgsNames, int cNames, int lcid, [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);

        [PreserveSig]
        int Invoke(int dispIdMember, ref Guid riid, [MarshalAs(UnmanagedType.U4)] int lcid, [MarshalAs(UnmanagedType.U4)] int dwFlags, ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out, MarshalAs(UnmanagedType.LPArray)] object[] pVarResult, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
    }

    #endregion

    /// <summary>
    /// Creation Factory for COMObject and derived types
    /// </summary>
    public static class Factory
    {
        #region Fields
        private static bool _initalized;
        private static bool _assemblyResolveEventConnected;
        private static List<COMObject> _globalObjectList = new List<COMObject>();
        private static List<IFactoryInfo> _factoryList = new List<IFactoryInfo>();
        private static Dictionary<string, Type> _proxyTypeCache = new Dictionary<string, Type>();
        private static Dictionary<string, Type> _wrapperTypeCache = new Dictionary<string, Type>();
        private static Dictionary<Guid, Guid> _hostCache = new Dictionary<Guid, Guid>();
        private static Dictionary<string, Dictionary<string, string>> _entitiesListCache = new Dictionary<string, Dictionary<string, string>>();
        private static Assembly _thisAssembly = Assembly.GetAssembly(typeof(COMObject));
        private static List<DependentAssembly> _dependentAssemblies = new List<DependentAssembly>();
        private static string[] _knownNetOfficeKeyTokens;

        private static object _factoryListLock = new object();
        private static object _comObjectLock = new object();
        private static object _globalObjectListLock = new object();

        #endregion

        #region Properties

        /// <summary>
        /// returns an array about currently loaded NetOfficeApi assemblies
        /// </summary>
        public static IFactoryInfo[] Assemblies
        {
            get
            {
                return _factoryList.ToArray();
            }
        }

        /// <summary>
        /// Returns count count of open proxies
        /// </summary>
        public static int ProxyCount
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
        /// <param name="proxyCount"></param>
        public delegate void ProxyCountChangedHandler(int proxyCount);

        /// <summary>
        /// notify info the count of proxies there open are changed
        /// in case of notify comes from event trigger created proxy the call comes from other thread
        /// </summary>
        public static event ProxyCountChangedHandler ProxyCountChanged;

        #endregion

        #region Factory Methods

        /// <summary>
        /// Must be called from client assembly for COMObject Support
        /// Recieve factory infos from all loaded NetOffice Assemblies in current application domain
        /// </summary>
        /// <param name="threadCulture">given value for Settings.ThreadCulture</param>
        /// <param name="enableThreadSafe">given value for Settings.EnableThreadSafe</param>
        public static void Initialize(System.Globalization.CultureInfo threadCulture, bool enableThreadSafe)
        {
            Settings.ThreadCulture = threadCulture;
            Settings.EnableThreadSafe = enableThreadSafe;
            Initialize();
        }

        /// <summary>
        /// Must be called from client assembly for COMObject Support
        /// Recieve factory infos from all loaded NetOffice Assemblies in current application domain
        /// </summary>
        /// <param name="threadCulture">given value for Settings.ThreadCulture</param>
        public static void Initialize(System.Globalization.CultureInfo threadCulture)
        {
            Settings.ThreadCulture = threadCulture;
            Initialize();
        }

        /// <summary>
        /// Must be called from client assembly for COMObject Support
        /// Recieve factory infos from all loaded NetOffice Assemblies in current application domain
        /// </summary>
        /// <param name="enableThreadSafe">given value for Settings.EnableThreadSafe</param>
        public static void Initialize(bool enableThreadSafe)
        {
            Settings.EnableThreadSafe = enableThreadSafe;
            Initialize();
        }

        /// <summary>
        /// Must be called from client assembly for COMObject Support
        /// Recieve factory infos from all loaded NetOffice Assemblies in current application domain
        /// </summary>
        /// <param name="cacheOptions">settings what NetOffice does with a previous existing cache(if exists)</param>
        public static void Initialize(CacheOptions cacheOptions)
        {
            Settings.CacheOptions = cacheOptions;
            Initialize();
        }

        /// <summary>
        /// Must be called from client assembly for COMObject Support
        /// Recieve factory infos from all loaded NetOfficeApi Assemblies in current application domain
        /// </summary>
        public static void Initialize()
        {
            _initalized = true;
            bool isLocked = false;
            try
            {
                if (Settings.EnableThreadSafe)
                {
                    Monitor.Enter(_factoryListLock);
                    isLocked = true;
                }

                DebugConsole.WriteLine("NetOffice.Factory.Initialize() DeepLevel:{0}", Settings.EnableDeepLoading);

                TryLoadAssembly("ExcelApi.dll");
                TryLoadAssembly("WordApi.dll");
                TryLoadAssembly("OutlookApi.dll");
                TryLoadAssembly("PowerPointApi.dll");
                TryLoadAssembly("AccessApi.dll");
                TryLoadAssembly("VisioApi.dll");
                TryLoadAssembly("MSProjectApi.dll");

                if (!_assemblyResolveEventConnected)
                {
                    AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
                    _assemblyResolveEventConnected = true;
                }
                
                ClearCache();
                AddNetOfficeAssemblies(Settings.EnableDeepLoading);
                AddDependentNetOfficeAssemblies();

                DebugConsole.WriteLine("Factory contains {0} assemblies", _factoryList.Count);
                DebugConsole.WriteLine("NetOffice.Factory.Initialize() passed");
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw (throwedException);
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
        private static void AddNetOfficeAssemblies(bool deepLevel)
        {
            _dependentAssemblies.Clear();

            if (deepLevel)
            {
                Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();
                foreach (Assembly domainAssembly in assemblies)
                {
                    foreach (AssemblyName itemName in domainAssembly.GetReferencedAssemblies())
                    {
                        if (ContainsNetOfficePublicKeyToken(itemName))
                        {
                            string assemblyName = itemName.Name;

                            DebugConsole.WriteLine(string.Format("Detect NetOffice assembly {0}.", assemblyName));
                            
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
                        DebugConsole.WriteLine(string.Format("Detect NetOffice assembly {0}.", assemblyName));

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
        private static void AddDependentNetOfficeAssemblies()
        {
            if (!Settings.EnableAdHocLoading)
                return;

            foreach (DependentAssembly dependAssembly in _dependentAssemblies)
            {
                if (!AssemblyExistsInFactoryList(dependAssembly.Name))
                {
                    string fileName = dependAssembly.ParentAssembly.CodeBase.Substring(0, dependAssembly.ParentAssembly.CodeBase.LastIndexOf("/")) + "/" + dependAssembly.Name;
                    fileName = fileName.Replace("/", "\\").Substring(8);

                    DebugConsole.WriteLine(string.Format("Try to load dependent assembly {0}.", fileName));

                    if (System.IO.File.Exists(fileName))
                    {
                        try
                        {
                            Assembly asssembly = Assembly.LoadFile(fileName);
                            AddAssembly(asssembly.GetName().Name, asssembly);
                        }
                        catch (Exception exception)
                        {
                            DebugConsole.WriteException(exception);
                        }
                    }
                    else
                    {
                        DebugConsole.WriteLine(string.Format("Assembly {0} not found.", fileName));
                    }
                }
            }
        }

        /// <summary>
        /// clears proxy/type/wrapper/assembly cache etc.
        /// </summary>
        private static void ClearCache()
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
        internal static void CheckInitialize()
        {
            if (!_initalized)
                Initialize();
        }

        /// <summary>
        /// clears factory informations List
        /// </summary>
        public static void ClearFactoryInformations()
        {
            bool isLocked = false;
            try
            {
                if (Settings.EnableThreadSafe)
                {
                    Monitor.Enter(_factoryListLock);
                    isLocked = true;
                }

                _factoryList.Clear();
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw (throwedException);
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
        internal static Dictionary<string, string> GetSupportedEntities(object comProxy)
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
        /// <param name="caller"></param>
        /// <param name="comProxy"></param>
        /// <param name="wrapperClassType"></param>
        /// <returns></returns>
        public static COMObject CreateKnownObjectFromComProxy(COMObject caller, object comProxy, Type wrapperClassType)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                if (null == comProxy)
                    return null;

                if (Settings.EnableThreadSafe)
                {
                    Monitor.Enter(_comObjectLock);
                    isLocked = true;
                }

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
                DebugConsole.WriteException(throwedException);
                throw throwedException;
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
        /// <param name="caller"></param>
        /// <param name="comProxyArray"></param>
        /// <param name="wrapperClassType"></param>
        /// <returns></returns>
        public static COMObject[] CreateKnownObjectArrayFromComProxy(COMObject caller, object[] comProxyArray, Type wrapperClassType)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                if (null == comProxyArray)
                    return null;

                if (Settings.EnableThreadSafe)
                {
                    Monitor.Enter(_comObjectLock);
                    isLocked = true;
                }

                Type comVariantType = null;
                COMObject[] newVariantArray = new COMObject[comProxyArray.Length];
                for (int i = 0; i < comProxyArray.Length; i++)
                    newVariantArray[i] = Activator.CreateInstance(wrapperClassType, new object[] { caller, comProxyArray[i], comVariantType }) as COMObject;

                return newVariantArray;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw throwedException;
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
        /// <returns>corresponding Wrapper class Instance or plain COMObject</returns>
        public static COMObject CreateObjectFromComProxy(COMObject caller, object comProxy)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                if (null == comProxy)
                    return null;

                if (Settings.EnableThreadSafe)
                {
                    Monitor.Enter(_comObjectLock);
                    isLocked = true;
                }

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
                DebugConsole.WriteException(throwedException);
                throw throwedException;
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
        public static COMObject CreateObjectFromComProxy(COMObject caller, object comProxy, Type comProxyType)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                if (null == comProxy)
                    return null;

                if (Settings.EnableThreadSafe)
                {
                    Monitor.Enter(_comObjectLock);
                    isLocked = true;
                }

                IFactoryInfo factoryInfo = GetFactoryInfo(comProxy);

                string className = TypeDescriptor.GetClassName(comProxy);
                string fullClassName = factoryInfo.AssemblyNamespace + "." + className;

                // create new classType
                COMObject newObject = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType, className, fullClassName);
                return newObject;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw throwedException;
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
        public static COMObject CreateObjectFromComProxy(IFactoryInfo factoryInfo, COMObject caller, object comProxy, Type comProxyType, string className, string fullClassName)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                if (Settings.EnableThreadSafe)
                {
                    Monitor.Enter(_comObjectLock);
                    isLocked = true;
                }

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
                DebugConsole.WriteException(throwedException);
                throw throwedException;
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
        public static COMObject[] CreateObjectArrayFromComProxy(COMObject caller, object[] comProxyArray)
        {
            CheckInitialize();
            bool isLocked = false;
            try
            {
                if (null == comProxyArray)
                    return null;

                if (Settings.EnableThreadSafe)
                {
                    Monitor.Enter(_comObjectLock);
                    isLocked = true;
                }

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
                DebugConsole.WriteException(throwedException);
                throw throwedException;
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

        public static void DisposeAllCOMProxies()
        {
            while (_globalObjectList.Count > 0)
                _globalObjectList[0].Dispose();
        }

        /// <summary>
        /// add object to global list
        /// </summary>
        /// <param name="proxy"></param>
        internal static void AddObjectToList(COMObject proxy)
        {
            bool isLocked = false;
            try
            {
                if (Settings.EnableThreadSafe)
                {
                    Monitor.Enter(_globalObjectList);
                    isLocked = true;
                }

                _globalObjectList.Add(proxy);

                if (null != ProxyCountChanged)
                    ProxyCountChanged(_globalObjectList.Count);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
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
        /// <param name="proxy"></param>
        internal static void RemoveObjectFromList(COMObject proxy)
        {
            bool isLocked = false;
            try
            {
                if (Settings.EnableThreadSafe)
                {
                    Monitor.Enter(_globalObjectList);
                    isLocked = true;
                }

                _globalObjectList.Remove(proxy);

                if (null != ProxyCountChanged)
                    ProxyCountChanged(_globalObjectList.Count);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
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
        /// <param name="itemAssembly"></param>
        /// <returns></returns>
        private static bool ContainsNetOfficeAttribute(Assembly itemAssembly)
        {
            try
            {
                List<string> dependAssemblies = new List<string>();
                object[] attributes = itemAssembly.GetCustomAttributes(true);
                foreach (object itemAttribute in attributes)
                {
                    string fullnameAttribute = itemAttribute.GetType().FullName;
                    if (fullnameAttribute == "NetOffice.NetOfficeAssemblyAttribute")
                        return true;
                }
                return false;
            }
            catch (System.IO.FileNotFoundException exception)
            {
                DebugConsole.WriteException(exception);
                return false;
            }
        }

        /// <summary>
        /// returns info the assembly is a NetOffice Api Assembly
        /// </summary>
        /// <param name="itemName"></param>
        /// <returns></returns>
        private static bool ContainsNetOfficePublicKeyToken(AssemblyName itemName)
        {
            try
            {
                string targetKeyToken = itemName.FullName.Substring(itemName.FullName.LastIndexOf(" ") +1);
                foreach (string item in KnownNetOfficeKeyTokens)
                {
                    if (item.EndsWith(targetKeyToken, StringComparison.InvariantCultureIgnoreCase))
                        return true;
                }
                return false;
            }
            catch (System.IO.FileNotFoundException exception)
            {
                DebugConsole.WriteException(exception);
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
                    Type thisType =  typeof(Factory);
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
        /// <param name="name"></param>
        /// <returns></returns>
        private static bool AssemblyExistsInFactoryList(string name)
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
        /// <param name="name"></param>
        /// <param name="itemAssembly"></param>
        /// <returns>list of dependend assemblies</returns>
        private static string[] AddAssembly(string name, Assembly itemAssembly)
        {
            List<string> dependAssemblies = new List<string>();
            object[] attributes = itemAssembly.GetCustomAttributes(true);
            foreach (object itemAttribute in attributes)
            {
                string fullnameAttribute = itemAttribute.GetType().FullName;
                if (fullnameAttribute == "NetOffice.NetOfficeAssemblyAttribute")
                {
                    Type factoryInfoType = itemAssembly.GetType(name + ".Utils.ProjectInfo");
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
                        DebugConsole.WriteLine("Recieve IFactoryInfo:{0}:{1}", factoryInfo.Assembly.FullName, factoryInfo.Assembly.FullName);
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
        /// <param name="typeInfo"></param>
        /// <returns></returns>
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
        /// <param name="comProxy"></param>
        /// <returns></returns>
        private static Guid GetParentLibraryGuid(object comProxy)
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
        /// <param name="comProxy"></param>
        /// <returns></returns>
        private static IFactoryInfo GetFactoryInfo(object comProxy)
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
                if (true == hostGuid.Equals(item.ComponentGuid))
                    return item;
            }

            // failback
            foreach (IFactoryInfo item in _factoryList)
            {
                if (item.Contains(className))
                    return item;
            }

            string message = string.Format("class {0}:{1} not found in loaded NetOffice Assemblies{2}", hostGuid, className, Environment.NewLine);
            message += string.Format("Currently loaded NetOfficeApi Assemblies{0}", Environment.NewLine);
            foreach (IFactoryInfo item in _factoryList)
                message += string.Format("Loaded NetOffice Assembly:{0} {1}{2}", item.ComponentGuid, item.Assembly.FullName, Environment.NewLine);

            throw new NetOfficeException(message);
        }

        /// <summary>
        /// AssemblyResolver Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                string directoryName = _thisAssembly.CodeBase.Substring(0, _thisAssembly.CodeBase.LastIndexOf("/"));
                directoryName = directoryName.Replace("/", "\\").Substring(8);
                string fileName = args.Name.Substring(0, args.Name.IndexOf(","));
                string fullFileName = System.IO.Path.Combine(directoryName, fileName + ".dll");
                if (System.IO.File.Exists(fullFileName))
                {
                    DebugConsole.WriteLine(string.Format("Try to resolve assembly {0}", args.Name));
                    Assembly assembly = System.Reflection.Assembly.Load(args.Name);
                    return assembly;
                }
                else
                {
                    DebugConsole.WriteLine(string.Format("Unable to resolve assembly {0}. The file file doesnt exists in current codebase.", args.Name));
                    return null;
                }
            }
            catch(Exception exception)
            {
                DebugConsole.WriteException(exception);
                return null;
            }
        }

        /// <summary>
        /// Assembly loader for multitargeting(host) scenarios
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static Assembly TryLoadAssembly(string fileName)
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
                        DebugConsole.WriteLine("Recieve IFactoryInfo:{0}:{1}", factoryInfo.Assembly.FullName, factoryInfo.Assembly.FullName);
                    }
                    return assembly;
                }
                else
                {
                    DebugConsole.WriteLine(string.Format("Unable to resolve assembly {0}. The assembly doesnt exists in current codebase.", fileName));
                    return null;
                }
            }
            catch (Exception exception)
            {
                DebugConsole.WriteException(exception);
                return null;
            }
        }

        #endregion

        #region Type

        /// <summary>
        /// returns the Type for comProxy or null if param not set
        /// </summary>
        /// <param name="comProxy"></param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type GetObjectType(object comProxy)
        {
            CheckInitialize();

            if (null == comProxy)
                return null;
            else
            {
                IFactoryInfo factoryInfo = GetFactoryInfo(comProxy);
                string className = TypeDescriptor.GetClassName(comProxy);
                string fullClassName = factoryInfo.AssemblyNamespace + "." + className;
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