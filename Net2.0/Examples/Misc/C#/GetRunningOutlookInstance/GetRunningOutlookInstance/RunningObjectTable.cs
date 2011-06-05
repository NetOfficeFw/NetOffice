using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Collections.Generic;
using System.Text;

namespace GetRunningOutlookInstance
{
    /// <summary>
    /// static helper class for accessing Running Object Table 
    /// </summary>
    /// <remarks>taken(and modified) from http://dotnet-snippets.de/dns/laufende-com-objekte-abfragen-SID526.aspx</remarks>
    public class RunningObjectTable
    {
        // Win32-API-call for reading ROT
        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable pprot);

        // Win32-API-call to create binding
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);


        /// <summary>
        /// returns native com proxy from power Point application object in Running Object Table
        /// </summary>
        /// <returns></returns>
        public static object GetRunningPowerPointInstanceFromROT()
        {
            return GetApplicationInstannceFromROT("Microsoft PowerPoint", "Application");
        }

        /// <summary>
        /// returns native com proxy from outlook application object in Running Object Table
        /// </summary>
        /// <returns></returns>
        public static object GetRunningOutlookInstanceFromROT()
        {
            return GetApplicationInstannceFromROT("Outlook", "Application");
        }

        private static object GetApplicationInstannceFromROT(string componentName, string className)
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
                    // create binding object
                    IBindCtx bindInfo;
                    CreateBindCtx(0, out bindInfo);

                    // query com proxy info       
                    object comInstance = null;
                    runningObjectTable.GetObject(monikerContainer[0], out comInstance);
                    string name = TypeDescriptor.GetClassName(comInstance);
                    string component = TypeDescriptor.GetComponentName(comInstance, false);
                    if ((component == componentName) && (name == className))
                    {
                        Marshal.ReleaseComObject(bindInfo);
                        return comInstance;
                    }
                    else
                        Marshal.ReleaseComObject(comInstance);

                    Marshal.ReleaseComObject(bindInfo);
                }

                // not running 
                return null;
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
