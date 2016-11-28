using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using COMTypes = System.Runtime.InteropServices.ComTypes;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Misc;

namespace NetOffice
{
    /// <summary>
    /// Try to find COM proxies from Desktop Subystem
    /// </summary>
    public static partial class RunningWindowTable
    {         
        #region Fields
 
        private static int _mainWindowTimeoutMilliseconds = 5000;
        private static int _childWindowTimeoutMilliseconds = 5000;
        private static Dictionary<string, AccessibleWindowTarget> Targets { get; set; }

        #endregion

        #region Ctor

        static RunningWindowTable()
        {
            InitializeTargets();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Timeout for the main window lookup in miliiseconds.
        /// Default: 5000, Possible Range 1000-90000
        /// </summary>
        public static int MainWindowTimeoutMilliseconds
        {
            get
            {
                return _mainWindowTimeoutMilliseconds;
            }
            set
            {
                if (value < 1000)
                    value = 1000;
                if (value > 90000)
                    value = 90000;
                _mainWindowTimeoutMilliseconds = value;
            }
        }

        /// <summary>
        /// Timeout for the child window lookup in miliiseconds.
        /// Default: 5000, Possible Range 1000-90000
        /// </summary>
        public static int ChildWindowTimeoutMilliseconds
        {
            get
            {
                return _childWindowTimeoutMilliseconds;
            }
            set
            {
                if (value < 1000)
                    value = 1000;
                if (value > 90000)
                    value = 90000;
                _childWindowTimeoutMilliseconds = value;
            }
        }
          
        #endregion

        #region Methods

        private static void InitializeTargets()
        {
            Dictionary<string, AccessibleWindowTarget> targets = new Dictionary<string, AccessibleWindowTarget>();

            AccessibleWindowTarget excel = new AccessibleWindowTarget("XLMAIN", new string[] { "XLDESK", "EXCEL7" });
            excel.AccPropertyName = "Application";
            targets.Add("Excel.Application", excel);

            AccessibleWindowTarget word = new AccessibleWindowTarget("OpusApp", new string[] { "_WwG" });
            word.AccPropertyName = "Application";
            targets.Add("Word.Application", word);

            AccessibleWindowTarget ppoint = new AccessibleWindowTarget("PP","FrameClass", NameCompareKind.StartEndWith, new string[] { "mdiClass"});
            ppoint.AccPropertyName = "Application";
            targets.Add("PowerPoint.Application", ppoint);

            Targets = targets;
        }

        /// <summary>
        /// Returns an accessible com proxy through the IAccessible interface
        /// </summary>
        /// <param name="type">target proxy type</param>
        /// <param name="throwExceptionIfNothingFound">throw an exception if no proxy found</param>
        /// <returns>com proxy instance or null</returns>
        public static object GetAccessibleProxy(ProxyType type, bool throwExceptionIfNothingFound)
        {
            IEnumerable<AccessibleWindowTarget> targets = ConvertToTargets(type);
            IDisposableEnumeration<ProxyInformation> result = GetAccessibleProxiesFromPath(targets, 1);
            if (result.Count == 0 && throwExceptionIfNothingFound)
                throw new ArgumentOutOfRangeException("type", "Unable to find accessible proxy");
            return result;
        }

        /// <summary>
        /// Returns all accessible com proxies through the IAccessible interface
        /// </summary>
        /// <param name="type">target proxy type</param>
        public static IDisposableEnumeration GetAccessibleProxies(ProxyType type)
        {
            IEnumerable<AccessibleWindowTarget> targets = ConvertToTargets(type);
            IDisposableEnumeration<ProxyInformation> result = GetAccessibleProxiesFromPath(targets);
            List<object> newResult = new List<object>();
            foreach (ProxyInformation item in result)
                newResult.Add(item.Proxy);
            return null;
        }

        /// <summary>
        ///  Returns all accessible com proxies and additional informations through the IAccessible interface 
        /// </summary>
        /// <param name="type">target proxy type</param>
        /// <returns>proxy information enumerator</returns>
        public static IDisposableEnumeration<ProxyInformation> GetAccessibleProxyInformations(ProxyType type)
        {
            IEnumerable<AccessibleWindowTarget> targets = ConvertToTargets(type);
            return GetAccessibleProxiesFromPath(targets);
        }

        /// <summary>
        /// Performs a lookup for window/child windows there implement the IAccessible interface to get a COM proxy
        /// </summary>
        /// <param name="targets">one ore more targets to lookup</param>
        /// <returns>proxy information ernumeration</returns>
        public static IDisposableEnumeration<ProxyInformation> GetAccessibleProxiesFromPath(IEnumerable<AccessibleWindowTarget> targets)
        {
            return GetAccessibleProxiesFromPath(targets, Int16.MaxValue);
        }
         
        /// <summary>
        /// Performs a lookup for window/child windows there implement the IAccessible interface to get a COM proxy
        /// </summary>
        /// <param name="targets">one ore more targets to lookup</param>
        /// <param name="maximumResultCount">maximum result count - the method abort if reached</param>
        /// <returns>proxy information ernumeration</returns>
        public static IDisposableEnumeration<ProxyInformation> GetAccessibleProxiesFromPath(IEnumerable<AccessibleWindowTarget> targets, int maximumResultCount)
        {
            if (null == targets || GetTargetsCount(targets) == 0)
            {
                return GetUnknownAccessibleProxiesFromPath(maximumResultCount);
            }
            else
                return GetKnownAccessibleProxiesFromPath(targets, maximumResultCount);
        }
         
        /// <summary>
        ///  Returns the count of accessible com proxies
        /// </summary>
        /// <param name="type">target proxy type</param>
        /// <returns>count of accessible proxies</returns>
        public static int GetAccessibleProxyCount(ProxyType type)
        {
            IEnumerable<AccessibleWindowTarget> targets = ConvertToTargets(type);
            IDisposableEnumeration<ProxyInformation> result = GetAccessibleProxiesFromPath(targets);
            int count = result.Count;
            result.Dispose();
            return count;
        }

        internal static Guid GetTypeGuid(COMTypes.ITypeInfo typeInfo)
        {
            if (null == typeInfo)
                return Guid.Empty;
            IntPtr attribPtr = IntPtr.Zero;
            typeInfo.GetTypeAttr(out attribPtr);
            COMTypes.TYPEATTR Attributes = (COMTypes.TYPEATTR)Marshal.PtrToStructure(attribPtr, typeof(COMTypes.TYPEATTR));
            Guid typeGuid = Attributes.guid;
            typeInfo.ReleaseTypeAttr(attribPtr);
            return typeGuid;
        }

        private static IEnumerable<AccessibleWindowTarget> ConvertToTargets(ProxyType type)
        {
            List<AccessibleWindowTarget> result = new List<AccessibleWindowTarget>();

            switch (type)
            {
                case ProxyType.ExcelApplication:
                    result.Add(Targets["Excel.Application"]);
                    break;
                case ProxyType.WordApplication:
                    result.Add(Targets["Excel.Application"]);
                    break;
                case ProxyType.PowerPointApplication:
                    result.Add(Targets["PowerPoint.Application"]);
                    break;
                case ProxyType.AllSupportedOfficeApplications:
                    result.Add(Targets["Excel.Application"]);
                    result.Add(Targets["Word.Application"]);
                    result.Add(Targets["PowerPoint.Application"]);
                    break;
                case ProxyType.All:
                    break;
                default:
                    throw new IndexOutOfRangeException();
            }
            return result;
        }

        // don't laughing for missing Linq knowledge because we still want to compile the core in .Net 2.0 
        private static int GetTargetsCount(IEnumerable<AccessibleWindowTarget> targets)
        {   
            if (null == targets)
                return 0;

            List<AccessibleWindowTarget> targetsImplementation = targets as List<AccessibleWindowTarget>;
            if (null != targetsImplementation)
            {
                return targetsImplementation.Count;
            }
            else
            { 
                int result = 0;
                foreach (AccessibleWindowTarget item in targets)
                    result++;
                return result;
            }
        }

        private static void EnumChildWindows(RunningWindowTableItemCollection list, IntPtr mainHandle, IntPtr handle, int maximumResultCount)
        {
            Tools.WndUtils.ChildWindowEnumerator childEnumerator =
                        new Tools.WndUtils.ChildWindowEnumerator(handle);
            IntPtr[] childWindows = childEnumerator.EnumerateWindows(10000);
            if (null != childWindows)
            {
                foreach (IntPtr item in childWindows)
                    DoAccTest(list, mainHandle, item, maximumResultCount);
            }
        }

        private static bool DoAccTest(RunningWindowTableItemCollection list, IntPtr mainHandle, IntPtr childHandle, int maximumResultCount)
        {
            object accObject = Tools.WndUtils.Win32.AccessibleObjectFromWindow(childHandle);
            if (null != accObject)
            {
                string name = TypeDescriptor.GetClassName(accObject);
                string component = TypeDescriptor.GetComponentName(accObject);
                string className = Tools.WndUtils.Win32.GetClassName(childHandle);
                IntPtr processID = Tools.WndUtils.Win32.GetWindowThreadProcessId(childHandle);
                ProxyInformation.ProcessElevation elevation = Tools.WndUtils.ProcessElevation.ConvertToProcessElevation(Tools.WndUtils.ProcessElevation.IsProcessElevated(processID));
                COMTypes.ITypeInfo typeInfo = RunningObjectTable.TryCreateTypeInfo(accObject);
                string id = GetTypeGuid(typeInfo).ToString();
                string libraryID = RunningObjectTable.GetParentLibraryGuid(typeInfo).ToString();
                if (null != typeInfo)
                    RunningObjectTable.ReleaseTypeInfo(typeInfo);
                ProxyInformation item =
                    new ProxyInformation(accObject, String.Format("{0}-{1}", childHandle, className), id, name, component, libraryID, processID, elevation);
                if(list.Count <= maximumResultCount)
                    list.Add(item);
                return true;
            }
            else
                return false;
        }

        private static IDisposableEnumeration<ProxyInformation> GetUnknownAccessibleProxiesFromPath(int maximumResultCount)
        {
            RunningWindowTableItemCollection result = new RunningWindowTableItemCollection();

            Tools.WndUtils.WindowEnumerator enumerator =
                     new Tools.WndUtils.WindowEnumerator(String.Empty);
            IntPtr[] mainWindows = enumerator.EnumerateWindows(10000);
            if (null != mainWindows)
            {
                foreach (IntPtr item in mainWindows)
                {
                    DoAccTest(result, item, item, maximumResultCount);
                    EnumChildWindows(result, item, item, maximumResultCount);
                    if (result.Count >= maximumResultCount)
                        break;
                }
            }

            return result;
        }

        private static IDisposableEnumeration<ProxyInformation> GetKnownAccessibleProxiesFromPath(IEnumerable<AccessibleWindowTarget> targets, int maximumResultCount)
        {
            if (null == targets)
                throw new ArgumentNullException("targets");

            RunningWindowTableItemCollection result = new RunningWindowTableItemCollection();

            if (maximumResultCount <= 0)
                return result;

            foreach (AccessibleWindowTarget target in targets)
            {
                Tools.WndUtils.WindowEnumerator enumerator =
                    new Tools.WndUtils.WindowEnumerator(
                        target.MainClassName, target.MainClassNameEnd, (Tools.WndUtils.WindowEnumerator.FilterMode)Convert.ToInt32(target.NameCompare));
                IntPtr[] mainHandles = enumerator.EnumerateWindows(_mainWindowTimeoutMilliseconds);
                if (null == mainHandles)
                    continue;
                foreach (IntPtr item in mainHandles)
                {
                    Tools.WndUtils.ChildWindowBatchEnumerator childEnumerator =
                        new Tools.WndUtils.ChildWindowBatchEnumerator(item);

                    foreach (string subItem in target.ChildPath)
                    {
                        childEnumerator.SearchOrder.Add(
                            new Tools.WndUtils.ChildWindowBatchEnumerator.SearchCriteria(subItem));
                    }
                    IntPtr[] childHandles = childEnumerator.EnumerateWindows(_childWindowTimeoutMilliseconds);
                    if (null == childHandles)
                        continue;

                    foreach (IntPtr childHandle in childHandles)
                    {
                        object accObject = Tools.WndUtils.Win32.AccessibleObjectFromWindow(childHandle);
                        if (null != accObject && accObject is MarshalByRefObject)
                        {
                            object targetProxy = null;
                            if (!String.IsNullOrEmpty(target.AccPropertyName))
                            { 
                                targetProxy = TryInvokeProperty(accObject, target.AccPropertyName);
                                Marshal.ReleaseComObject(accObject);
                            }
                            else
                                targetProxy = accObject;

                            if (null != targetProxy)
                            {
                                string itemComponentName = TypeDescriptor.GetComponentName(targetProxy);
                                COMTypes.ITypeInfo typeInfo = RunningObjectTable.TryCreateTypeInfo(targetProxy);
                                string library = RunningObjectTable.GetParentLibraryGuid(typeInfo).ToString();
                                string id = GetTypeGuid(typeInfo).ToString();
                                string itemClassName = TypeDescriptor.GetClassName(targetProxy);
                                string itemCaption = itemClassName;
                                if (!String.IsNullOrWhiteSpace(itemClassName) && !String.IsNullOrWhiteSpace(itemComponentName))
                                    itemCaption = String.Format("{0} {1}", itemComponentName, itemClassName);

                                IntPtr procID = Tools.WndUtils.Win32.GetWindowThreadProcessId(childHandle);
                                ProxyInformation.ProcessElevation procElevation =
                                   Tools.WndUtils.ProcessElevation.ConvertToProcessElevation(Tools.WndUtils.ProcessElevation.IsProcessElevated(procID));
                                
                                ProxyInformation info = new ProxyInformation(targetProxy,
                                    itemCaption, id, itemClassName, itemComponentName, library, procID, procElevation);

                                result.Add(info);
                                if (null != typeInfo)
                                    RunningObjectTable.ReleaseTypeInfo(typeInfo);

                                if (result.Count >= maximumResultCount)
                                    return result;
                            }
                        }
                    }
                }
            }

            return result;
        }

        private static object TryInvokeProperty(object proxy, string propertyName)
        {
            if (null == proxy || String.IsNullOrEmpty(propertyName))
                return null;

            try
            {
                return proxy.GetType().InvokeMember(propertyName,
                    BindingFlags.GetProperty, null, proxy, new object[0]);
            }
            catch
            {
                return null;
            }
        }

        #endregion
    }
}
