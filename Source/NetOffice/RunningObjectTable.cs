using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// ROT Wrapper
    /// </summary>
    public static class RunningObjectTable
    {
        /// <summary>
        /// some office applications in specific version use the "Microsoft" prefix in the COM server name
        /// </summary>
        private static readonly string _ballmersPlace = "Microsoft ";

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable pprot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);

        /// <summary>
        /// returns a running com proxy from the running object table. the method takes the first proxy there matched with the input parameters
        /// </summary>
        /// <param name="componentName">component name, for example Excel</param>
        /// <param name="className">class name, for example Application</param>
        /// <param name="throwOnError">throw an exception if no proxy was found</param>
        /// <returns>a native COM proxy</returns>
        public static object GetActiveProxyFromROT(string componentName, string className, bool throwOnError)
        {
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
                        componentNameEqual = ((_ballmersPlace + componentName).Equals(component, StringComparison.InvariantCultureIgnoreCase));
                        if (componentNameEqual && classNameEqual)
                        {
                            return comInstance;
                        }
                        else
                        {
                            if (comInstance.GetType().IsCOMObject)
                                Marshal.ReleaseComObject(comInstance);
                        }
                    }
                }

                if (throwOnError)
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
                    Marshal.ReleaseComObject(runningObjectTable);
                if (monikerList != null)
                    Marshal.ReleaseComObject(monikerList);
            }
        }

        /// <summary>
        /// returns all running com proxies from the running object table there matched with the input parameters 
        /// </summary>
        /// <param name="componentName">component name, for example Excel</param>
        /// <param name="className">class name, for example Application</param>
        /// <returns>COM proxy list</returns>
        public static List<object> GetActiveProxiesFromROT(string componentName, string className)
        {
            IEnumMoniker monikerList = null;
            IRunningObjectTable runningObjectTable = null;
            List<object> resultList = new List<object>();
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

                    // get class name and component name
                    string name = TypeDescriptor.GetClassName(comInstance);
                    string component = TypeDescriptor.GetComponentName(comInstance, false);

                    // match for equal and add to list
                    bool componentNameEqual = (componentName.Equals(component, StringComparison.InvariantCultureIgnoreCase));
                    bool classNameEqual = (className.Equals(name, StringComparison.InvariantCultureIgnoreCase));

                    if (componentNameEqual && classNameEqual)
                    {
                        resultList.Add(comInstance);
                    }
                    else
                    {
                        componentNameEqual = ((_ballmersPlace + componentName).Equals(component, StringComparison.InvariantCultureIgnoreCase));
                        if (componentNameEqual && classNameEqual)
                        {
                            resultList.Add(comInstance);
                        }
                        else
                        {
                            if (comInstance.GetType().IsCOMObject)
                                Marshal.ReleaseComObject(comInstance);
                        }

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
                    Marshal.ReleaseComObject(runningObjectTable);
                if (monikerList != null)
                    Marshal.ReleaseComObject(monikerList);
            }
        }
    }
}