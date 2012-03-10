using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using COMTypes = System.Runtime.InteropServices.ComTypes;

namespace LateBindingApi.Core
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

        private static List<COMObject> _globalObjectList = new List<COMObject>();
        private static List<IFactoryInfo> _factoryList = new List<IFactoryInfo>();
        private static Dictionary<string, Type> _proxyTypeCache = new Dictionary<string, Type>();
        private static Dictionary<string, Type> _wrapperTypeCache = new Dictionary<string, Type>();
        private static Dictionary<Guid, Guid> _hostCache = new Dictionary<Guid, Guid>();

        #endregion

        #region Properties

        /// <summary>
        /// returns an array about currently loaded LateBindingApi assemblies
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
        /// Recieve FactoryInfos from all loaded LateBindingApi based Assemblies
        /// </summary>
        public static void Initialize()
        {
            try
            {
                DebugConsole.WriteLine("LateBindingApi.Core.Factory.Initialize()");

                List<string> dependAssemblies = new List<string>();
                Assembly callingAssembly = System.Reflection.Assembly.GetCallingAssembly();
                foreach (AssemblyName item in callingAssembly.GetReferencedAssemblies())
                {
                    DebugConsole.WriteLine(string.Format("Load assembly {0}.", item.Name));

                    Assembly itemAssembly = Assembly.Load(item);
                    string[] depends = AddAssembly(item.Name, itemAssembly);
                    foreach (string depend in depends)
                    {
                        bool found = false;
                        foreach (string itemExistingDependency in dependAssemblies)
                        {
                            if (depend == itemExistingDependency)
                            {
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                            dependAssemblies.Add(depend);
                    }
                }
                
                // try load non loaded dependent assemblies
                if (Settings.EnableAdHocLoading)
                { 
                    foreach (string itemAssemblyName in dependAssemblies)
                    {
                        DebugConsole.WriteLine(string.Format("Try to load dependent assembly {0}.", itemAssemblyName));

                        string fileName = callingAssembly.CodeBase.Substring(0, callingAssembly.CodeBase.LastIndexOf("/"))+ "/" + itemAssemblyName;
                        fileName = fileName.Replace("/", "\\").Substring(8);

                        if (System.IO.File.Exists(fileName))
                        {
                            try
                            {
                                Assembly dependAssembly = Assembly.LoadFile(fileName);
                                AddAssembly(dependAssembly.GetName().Name, dependAssembly);
                            }
                            catch (Exception exception)
                            {
                                DebugConsole.WriteException(exception);
                            }
                        }
                        else
                        {
                            DebugConsole.WriteLine(string.Format("Assembly {0} not found.", itemAssemblyName));
                        }
                    }
                }

                DebugConsole.WriteLine("LateBindingApi.Core.Factory.Initialize() passed");
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw (throwedException);
            }
        }

        /// <summary>
        /// clears factory informations List
        /// </summary>
        public static void Clear()
        {
            _factoryList.Clear();
        }

        /// <summary>
        /// creates an entity support list for a proxy
        /// </summary>
        /// <param name="comProxy"></param>
        /// <returns></returns>
        internal static Dictionary<string, string> GetSupportedEntities(object comProxy)
        {
            Dictionary<string, string> supportList = new Dictionary<string, string>();
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
            try
            {
                if (null == comProxy)
                    return null;

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
            try
            {
                if (null == comProxyArray)
                    return null;

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
        }

        /// <summary>
        /// creates a new COMObject based on classType of comProxy 
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxy">new created proxy</param>
        /// <returns>corresponding Wrapper class Instance or plain COMObject</returns>
        public static COMObject CreateObjectFromComProxy(COMObject caller, object comProxy)
        {
            try
            {
                if (null == comProxy)
                    return null;

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
            try
            {
                if (null == comProxy)
                    return null;

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
            try
            {
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
                    classType = factoryInfo.Assembly.GetType(fullClassName,false,true);
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
        }

        /// <summary>
        ///  creates a new COMObject array
        /// </summary>
        /// <param name="caller">parent there have created comProxy</param>
        /// <param name="comProxyArray">new created proxy array</param>
        /// <returns>corresponding Wrapper class Instance array or plain COMObject array</returns>
        public static COMObject[] CreateObjectArrayFromComProxy(COMObject caller, object[] comProxyArray)
        {
            try
            {
                if (null == comProxyArray)
                    return null;

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
        }

        #endregion

        #region Object List Methods

        /// <summary>
        /// dispose all open objects
        /// </summary>
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
            _globalObjectList.Add(proxy);

            if (null != ProxyCountChanged)
                ProxyCountChanged(_globalObjectList.Count);
        }

        /// <summary>
        /// remove object from global list
        /// </summary>
        /// <param name="proxy"></param>
        internal static void RemoveObjectFromList(COMObject proxy)
        {
            _globalObjectList.Remove(proxy);

            if (null != ProxyCountChanged)
                ProxyCountChanged(_globalObjectList.Count);
        }

        #endregion

        #region Private Methods
        
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
                if (fullnameAttribute == "LateBindingApi.Core.LateBindingAttribute")
                {
                    Type factoryInfoType = itemAssembly.GetType(name + ".Utils.ProjectInfo");
                    IFactoryInfo factoryInfo = Activator.CreateInstance(factoryInfoType) as IFactoryInfo;

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
                        _factoryList.Add(factoryInfo);

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
                string notInitMessage = "Factory are not initialized with LateBindingApi assemblies." + Environment.NewLine;
                notInitMessage = "Please call LateBindingApi.Core.Factory.Initialize()";
                throw new LateBindingApiException(notInitMessage);
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

            string message = string.Format("class {0}:{1} not found in loaded LateBindingApi Assemblies{2}", hostGuid, className, Environment.NewLine);
            message += string.Format("Currently loaded LateBindingApi Assemblies{0}", Environment.NewLine);
            foreach (IFactoryInfo item in _factoryList)
                message += string.Format("Loaded LateBindingApi Assembly:{0} {1}{2}", item.ComponentGuid, item.Assembly.FullName, Environment.NewLine);

            throw new LateBindingApiException(message);
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
