﻿using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using COMTypes = System.Runtime.InteropServices.ComTypes;
using NetOffice.CollectionsGeneric;
using NetOffice.Contribution.CollectionsGeneric;
using System.Collections.Generic;

namespace NetOffice.Running
{
    /// <summary>
    /// A Running Object Table(ROT) Wrapper
    /// </summary>
    public static class RunningObjectTable
    {
        #region Imports

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable pprot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);

        #endregion

        #region Nested
         
        internal class RunningObjectTableItemCollection : SortableBindingList<ProxyInformation>, IDisposableSequence<ProxyInformation>
        {
            #region IDisposableEnumeration<ProxyInformation>

            /// <summary>
            /// Dispose the instance
            /// </summary>
            public void Dispose()
            {
                foreach (object item in this)
                {
                    IDisposable disposeItem = item as IDisposable;
                    if (null != disposeItem)
                    {
                        try
                        {
                            disposeItem.Dispose();
                        }
                        catch
                        {
                            ;
                        }
                    }
                    if (item is MarshalByRefObject)
                    {
                        Marshal.ReleaseComObject(item);
                    }
                }
                Clear();
            }

            public void Dispose(ProxyInformation keepAliveItem)
            {
                foreach (object item in this)
                {
                    if (item == keepAliveItem)
                        continue;

                    IDisposable disposeItem = item as IDisposable;
                    if (null != disposeItem)
                    {
                        try
                        {
                            disposeItem.Dispose();
                        }
                        catch
                        {
                            ;
                        }
                    }
                    if (item is MarshalByRefObject)
                    {
                        Marshal.ReleaseComObject(item);
                    }
                }
                Clear();
            }

            #endregion

            #region Overrides

            /// <summary>
            /// Returns a System.String that represents the instance
            /// </summary>
            /// <returns>System.String</returns>
            public override string ToString()
            {
                return String.Format("{0} Items", Count);
            }

            #endregion
        }

        #endregion

        #region Fields

        /// <summary>
        /// Some MS Office applications in specific versions use the "Microsoft " prefix in the COM server name.
        /// </summary>
        private static readonly string ComponentNamePrefix = "Microsoft ";

        #endregion

        #region Methods

        /// <summary>
        /// Returns the count of open com proxies
        /// </summary>
        /// <param name="componentName">component name or null as wildcard</param>
        /// <param name="className">class name or null as wildcard</param>
        /// <returns>count of open com proxies</returns>
        public static int GetProxyCount(string componentName, string className)
        {
            int totalCount = 0;
            IEnumMoniker monikerList = null;
            IRunningObjectTable runningObjectTable = null;
            try
            {
                // query table and returns null if no objects runnings
                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null)
                    return totalCount;

                // query moniker & reset
                runningObjectTable.EnumRunning(out monikerList);
                monikerList.Reset();

                IMoniker[] monikerContainer = new IMoniker[1];
                IntPtr pointerFetchedMonikers = IntPtr.Zero;

                // fetch all moniker
                while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                {
                    // query com proxy info      
                    object comInstance = null;
                    runningObjectTable.GetObject(monikerContainer[0], out comInstance);
                    if (null == comInstance)
                        continue;

                    // get class name and component name
                    string name = TypeDescriptor.GetClassName(comInstance);
                    string component = TypeDescriptor.GetComponentName(comInstance, false);

                    // match for equal and return
                    bool componentNameEqual = String.IsNullOrEmpty(componentName) ? true :
                        (componentName.Equals(component, StringComparison.InvariantCultureIgnoreCase));

                    bool classNameEqual = String.IsNullOrEmpty(className) ? true :
                        (className.Equals(name, StringComparison.InvariantCultureIgnoreCase));

                    if (componentNameEqual && classNameEqual)
                    {
                        totalCount++;
                    }
                    else
                    {
                        componentNameEqual = ((ComponentNamePrefix + componentName).Equals(component, StringComparison.InvariantCultureIgnoreCase));
                        if (componentNameEqual && classNameEqual)
                        {
                            totalCount++;
                        }
                    }

                    if (comInstance.GetType().IsCOMObject)
                        TryMarshalReleaseComObject(comInstance);
                }

                return totalCount;
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }
            finally
            {
                // release proxies
                if (runningObjectTable != null)
                    TryMarshalReleaseComObject(runningObjectTable);
                if (monikerList != null)
                    TryMarshalReleaseComObject(monikerList);
            }
        }

        /// <summary>
        /// Returns a running com proxy from the running object table. the method takes the first proxy there matched with the input parameters.
        /// WARNING: the method returns always the first com proxy from the running object table if multiple (match) proxies exists.
        /// </summary>
        /// <param name="componentName">component name, for example Excel</param>
        /// <param name="className">class name, for example Application</param>
        /// <param name="throwExceptionIfNothingFound">throw an exception if no proxy was found</param>
        /// <returns>a native COM proxy</returns>
        public static object GetActiveProxy(string componentName, string className, bool throwExceptionIfNothingFound)
        {
            if (String.IsNullOrEmpty(componentName))
                throw new ArgumentNullException("componentName");
            if (String.IsNullOrEmpty(className))
                throw new ArgumentNullException("className");

            IEnumMoniker monikerList = null;
            IRunningObjectTable runningObjectTable = null;
            try
            {
                // query table and returns null if no objects runnings
                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null)
                    return null;

                // query moniker & reset
                runningObjectTable.EnumRunning(out monikerList);
                monikerList.Reset();

                IMoniker[] monikerContainer = new IMoniker[1];
                IntPtr pointerFetchedMonikers = IntPtr.Zero;

                // fetch all moniker
                while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                {
                    // query com proxy info      
                    object comInstance = null;
                    runningObjectTable.GetObject(monikerContainer[0], out comInstance);
                    if (null == comInstance)
                        continue;

                    // get class name and component name
                    string name = TypeDescriptor.GetClassName(comInstance);
                    string component = TypeDescriptor.GetComponentName(comInstance, false);

                    // match for equal and return
                    bool componentNameEqual = (componentName.Equals(component, StringComparison.InvariantCultureIgnoreCase));
                    bool classNameEqual = (className.Equals(name, StringComparison.InvariantCultureIgnoreCase));

                    if (componentNameEqual && classNameEqual)
                    {
                        return comInstance;
                    }
                    else
                    {
                        componentNameEqual = ((ComponentNamePrefix + componentName).Equals(component, StringComparison.InvariantCultureIgnoreCase));
                        if (componentNameEqual && classNameEqual)
                        {
                            return comInstance;
                        }
                        else
                        {
                            if (comInstance.GetType().IsCOMObject)
                                TryMarshalReleaseComObject(comInstance);
                        }
                    }
                }

                if (throwExceptionIfNothingFound)
                    throw new COMException("Target instance is not running.");
                else
                    return null;
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }
            finally
            {
                // release proxies
                if (runningObjectTable != null)
                    TryMarshalReleaseComObject(runningObjectTable);
                if (monikerList != null)
                    TryMarshalReleaseComObject(monikerList);
            }
        }

        /// <summary>
        /// Returns all running com proxies from the running object table there matched with the input parameters 
        /// WARNING: the method returns always the first com proxy from the running object table if multiple (match) proxies exists.
        /// </summary>
        /// <param name="componentName">component name, for example Excel, null is a wildcard </param>
        /// <param name="className">class name, for example Application, null is a wildcard </param>
        /// <returns>COM proxy enumerator</returns>
        public static IDisposableSequence GetActiveProxies(string componentName, string className)
        {          
            IEnumMoniker monikerList = null;
            IRunningObjectTable runningObjectTable = null;
            List<object> resultList = new List<object>();
            try
            {
                // query table and returns null if no objects running
                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null)
                    return null;

                // query moniker & reset
                runningObjectTable.EnumRunning(out monikerList);
                monikerList.Reset();

                IMoniker[] monikerContainer = new IMoniker[1];
                IntPtr pointerFetchedMonikers = IntPtr.Zero;

                // fetch all moniker
                while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                {
                    // query com proxy info      
                    object comInstance = null;
                    runningObjectTable.GetObject(monikerContainer[0], out comInstance);
                    if (null == comInstance)
                        continue;

                    // get class name and component name
                    string name = TypeDescriptor.GetClassName(comInstance);
                    string component = TypeDescriptor.GetComponentName(comInstance, false);

                    // match for equal and add to list
                    bool componentNameEqual = String.IsNullOrWhiteSpace(component) ? true : 
                        (componentName.Equals(component, StringComparison.InvariantCultureIgnoreCase));
                    bool classNameEqual = String.IsNullOrWhiteSpace(className) ? true : 
                        (className.Equals(name, StringComparison.InvariantCultureIgnoreCase));

                    if (componentNameEqual && classNameEqual)
                    {
                        resultList.Add(comInstance);
                    }
                    else
                    {
                        componentNameEqual = ((ComponentNamePrefix + componentName).Equals(component, StringComparison.InvariantCultureIgnoreCase));
                        if (componentNameEqual && classNameEqual)
                        {
                            resultList.Add(comInstance);
                        }
                        else
                        {
                            if (comInstance.GetType().IsCOMObject)
                                TryMarshalReleaseComObject(comInstance);
                        }
                    }
                }

                return new DisposableObjectList(resultList.ToArray());
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }
            finally
            {
                // release proxies
                if (runningObjectTable != null)
                    TryMarshalReleaseComObject(runningObjectTable);
                if (monikerList != null)
                    TryMarshalReleaseComObject(monikerList);
            }
        }

        /// <summary>
        /// Returns all running com proxies + add. information from the running object table there matched with the input parameters
        /// WARNING: the method returns always the first com proxy from the running object table if multiple (match) proxies exists.
        /// </summary>
        /// <param name="componentName">name of the target component</param>
        /// <param name="className">name of the target proxy class name</param>
        /// <returns>IDisposableEnumeration with proxy information</returns>
        public static IDisposableSequence<ProxyInformation> GetActiveProxyInformations(string componentName, string className)
        {
            IEnumMoniker monikerList = null;
            IRunningObjectTable runningObjectTable = null;
            RunningObjectTableItemCollection resultList = new RunningObjectTableItemCollection();
            try
            {
                // query table and returns null if no objects running
                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null)
                    return null;

                // query moniker & reset
                runningObjectTable.EnumRunning(out monikerList);
                monikerList.Reset();

                IMoniker[] monikerContainer = new IMoniker[1];
                IntPtr pointerFetchedMonikers = IntPtr.Zero;

                // fetch all moniker
                while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                {
                    // query com proxy info      
                    object comInstance = null;
                    runningObjectTable.GetObject(monikerContainer[0], out comInstance);
                    if (null == comInstance)
                        continue;

                    // match for equal and add to list
                    string name = TypeDescriptor.GetClassName(comInstance);
                    string component = TypeDescriptor.GetComponentName(comInstance, false);
                    bool componentNameEqual = String.IsNullOrWhiteSpace(componentName) ? true :
                        (componentName.Equals(component, StringComparison.InvariantCultureIgnoreCase));
                    bool classNameEqual = String.IsNullOrWhiteSpace(className) ? true :
                        (className.Equals(name, StringComparison.InvariantCultureIgnoreCase));
                    bool match = false;

                    if (componentNameEqual && classNameEqual)
                    {
                        match = true;
                    }
                    else
                    {
                        componentNameEqual = ((ComponentNamePrefix + componentName).Equals(component, StringComparison.InvariantCultureIgnoreCase));
                        if (componentNameEqual && classNameEqual)
                        {
                            match = true;
                        }
                        else
                        {
                            if (comInstance.GetType().IsCOMObject)
                                TryMarshalReleaseComObject(comInstance);
                        }
                    }

                    if (match)
                    {
                        IBindCtx bindInfo = null;
                        string displayName = String.Empty;
                        Guid classID = Guid.Empty;
                        if (CreateBindCtx(0, out bindInfo) == 0)
                        {
                            monikerContainer[0].GetDisplayName(bindInfo, null, out displayName);
                            monikerContainer[0].GetClassID(out classID);
                            TryMarshalReleaseComObject(bindInfo);
                             
                        }

                        string itemClassName = TypeDescriptor.GetClassName(comInstance);
                        string itemComponentName = TypeDescriptor.GetComponentName(comInstance);

                        COMTypes.ITypeInfo typeInfo = null;
                        string itemLibrary = String.Empty;
                        if (classID != Guid.Empty)
                        { 
                            typeInfo = TryCreateTypeInfo(comInstance);
                            itemLibrary = null != typeInfo ? GetParentLibraryGuid(typeInfo).ToString() : String.Empty;
                        }

                        string itemID = classID != Guid.Empty ? classID.ToString() : String.Empty;

                        ProxyInformation entry = 
                            new ProxyInformation(comInstance, displayName, itemID, itemClassName,
                            itemComponentName, itemLibrary, IntPtr.Zero, ProxyInformation.ProcessElevation.Unknown);

                        resultList.Add(entry);
                        if (classID != Guid.Empty && typeInfo != null)
                            ReleaseTypeInfo(typeInfo);
                    }
                }

                return resultList;
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }
            finally
            {
                // release proxies
                if (runningObjectTable != null)
                    TryMarshalReleaseComObject(runningObjectTable);
                if (monikerList != null)
                    TryMarshalReleaseComObject(monikerList);
            }
        }

        /// <summary>
        /// Returns all running com proxies + add. information from the running object table there matched with the input parameters
        /// WARNING: the method returns always the first com proxy from the running object table if multiple (match) proxies exists.
        /// </summary>
        /// <returns>IDisposableEnumeration with proxy information</returns>
        public static IDisposableSequence<ProxyInformation> GetActiveProxyInformations()
        {
            IEnumMoniker monikerList = null;
            IRunningObjectTable runningObjectTable = null;
            RunningObjectTableItemCollection resultList = new RunningObjectTableItemCollection();
            try
            {
                // query table and returns null if no objects running
                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null)
                    return null;

                // query moniker & reset
                runningObjectTable.EnumRunning(out monikerList);
                monikerList.Reset();

                IMoniker[] monikerContainer = new IMoniker[1];
                IntPtr pointerFetchedMonikers = IntPtr.Zero;

                // fetch all moniker
                while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                {
                    // query com proxy info      
                    object comInstance = null;
                    runningObjectTable.GetObject(monikerContainer[0], out comInstance);
                    if (null == comInstance)
                        continue;

                    string name = TypeDescriptor.GetClassName(comInstance);
                    string component = TypeDescriptor.GetComponentName(comInstance, false);                    

                    IBindCtx bindInfo = null;
                    string displayName = String.Empty;
                    Guid classID = Guid.Empty;
                    if (CreateBindCtx(0, out bindInfo) == 0)
                    {
                        monikerContainer[0].GetDisplayName(bindInfo, null, out displayName);
                        monikerContainer[0].GetClassID(out classID);
                        TryMarshalReleaseComObject(bindInfo);

                    }

                    string itemClassName = TypeDescriptor.GetClassName(comInstance);
                    string itemComponentName = TypeDescriptor.GetComponentName(comInstance);

                    COMTypes.ITypeInfo typeInfo = null;
                    string itemLibrary = String.Empty;
                    if (classID != Guid.Empty)
                    {
                        typeInfo = TryCreateTypeInfo(comInstance);
                        itemLibrary = null != typeInfo ? GetParentLibraryGuid(typeInfo).ToString() : String.Empty;
                    }

                    string itemID = classID != Guid.Empty ? classID.ToString() : String.Empty;

                    ProxyInformation entry =
                        new ProxyInformation(comInstance, displayName, itemID, itemClassName,
                        itemComponentName, itemLibrary, IntPtr.Zero, ProxyInformation.ProcessElevation.Unknown);

                    resultList.Add(entry);
                    if (classID != Guid.Empty && typeInfo != null)
                        ReleaseTypeInfo(typeInfo);
                }

                return resultList;
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }
            finally
            {
                // release proxies
                if (runningObjectTable != null)
                    TryMarshalReleaseComObject(runningObjectTable);
                if (monikerList != null)
                    TryMarshalReleaseComObject(monikerList);
            }
        }

        internal static void ReleaseTypeInfo(COMTypes.ITypeInfo typeInfo)
        {
            if(null != typeInfo)
                Marshal.ReleaseComObject(typeInfo);
        }

        internal static COMTypes.ITypeInfo TryCreateTypeInfo(object comProxy)
        {
            IDispatch dispatcher = comProxy as IDispatch;
            if (null == dispatcher)
                return null;
            try
            {
                COMTypes.ITypeInfo typeInfo = dispatcher.GetTypeInfo(0, 0);
                return typeInfo;
            }
            catch
            {
                // Seems to be check for null after cast to IDispatch is useless
                // because we got an InvalidCast exception from GetTypeInfo sometimes
                return null;
            }
        }

        internal static Guid GetParentLibraryGuid(COMTypes.ITypeInfo typeInfo)
        {
            if (null == typeInfo)
                return Guid.Empty;

            COMTypes.ITypeLib parentTypeLib = null;
            Guid parentGuid = Guid.Empty;

            int i = 0;
            typeInfo.GetContainingTypeLib(out parentTypeLib, out i);

            IntPtr attributesPointer = IntPtr.Zero;
            parentTypeLib.GetLibAttr(out attributesPointer);

            COMTypes.TYPELIBATTR attributes = (COMTypes.TYPELIBATTR)Marshal.PtrToStructure(attributesPointer, typeof(COMTypes.TYPELIBATTR));
            parentGuid = attributes.guid;
            parentTypeLib.ReleaseTLibAttr(attributesPointer);
            Marshal.ReleaseComObject(parentTypeLib);
             
            return parentGuid;
        }

        internal static bool TryMarshalReleaseComObject(object comProxy)
        {
            try
            {
                Marshal.ReleaseComObject(comProxy);
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion
    }
}