using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using NetOffice.Misc;

namespace NetOffice
{
    static partial class RunningWindowTable
    {
        #region Nested

        /// <summary>
        /// Determine which kind of proxy is wanted
        /// </summary>
        public enum ProxyType
        {
            /// <summary>
            /// Excel Application Proxy
            /// </summary>
            ExcelApplication = 0,

            /// <summary>
            /// Word Application Proxy
            /// </summary>
            WordApplication = 1,

            /// <summary>
            /// PowerPoint Application Proxy
            /// </summary>
            PowerPointApplication = 2,

            /// <summary>
            /// All supported Office applications together
            /// </summary>
            AllSupportedOfficeApplications = 3,

            /// <summary>
            /// All Together, incl. unkown
            /// </summary>
            All = 4
        }

        /// <summary>
        /// Determine how RunningWindowTable do compares the main window class name.
        /// Upper/Lower case is always ignored
        /// </summary>
        public enum NameCompareKind
        {
            /// <summary>
            /// Must match completely with MainClassName
            /// </summary>
            Equal = 0,

            /// <summary>
            /// Must start with MainClassName
            /// </summary>
            StartsWith = 1,

            /// <summary>
            /// Must ends with MainClassName
            /// </summary>
            EndsWith = 2,

            /// <summary>
            /// Must start with MainClassName and ends with MainClassNameEnd
            /// </summary>
            StartEndWith = 3
        }

        /// <summary>
        /// Path to a window or child window that implements the IAccessible interface
        /// </summary>
        public class AccessibleWindowTarget
        {
            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            public AccessibleWindowTarget()
            {
                ChildPath = new List<string>();
            }

            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="mainClassName">name of the main window class</param>
            public AccessibleWindowTarget(string mainClassName)
            {
                MainClassName = mainClassName;
                ChildPath = new List<string>();
            }

            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="mainClassName">name of the main window class</param>
            /// <param name="childPath">path to a child window</param>
            public AccessibleWindowTarget(string mainClassName, IEnumerable<string> childPath)
            {
                MainClassName = mainClassName;
                ChildPath = new List<string>();
                if (null != childPath)
                    ChildPath.AddRange(childPath);
            }

            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="mainClassName">name of the main window class</param>
            /// <param name="mainClassNameEnd">ends with part of main window class name if its compared as StartsEndWith</param>
            /// <param name="nameCompare">Determine how RunningWindowTable do compares the main window class name</param>
            /// <param name="childPath">path to a child window</param>
            public AccessibleWindowTarget(string mainClassName, string mainClassNameEnd, NameCompareKind nameCompare, IEnumerable<string> childPath)
            {
                MainClassName = mainClassName;
                MainClassNameEnd = mainClassNameEnd;
                NameCompare = nameCompare;
                ChildPath = new List<string>();
                if (null != childPath)
                    ChildPath.AddRange(childPath);
            }

            /// <summary>
            /// Name of the main window class
            /// </summary>
            public string MainClassName { get; set; }

            /// <summary>
            /// EndsWith Part of main window class name if its compared as StartsEndWith
            /// </summary>
            public string MainClassNameEnd { get; set; }

            /// <summary>
            /// Determine how RunningWindowTable do compares the main window class name
            /// </summary>
            public NameCompareKind NameCompare { get; set; }

            /// <summary>
            /// Path to a child window
            /// </summary>
            public List<string> ChildPath { get; private set; }

            /// <summary>
            /// If its not null or empty RunningObjectTable want invoke a property with these name from IAccessible proxy
            /// </summary>
            public string AccPropertyName { get; set; }

            /// <summary>
            /// Returns a System.String that represents the instance
            /// </summary>
            /// <returns>System.String</returns>
            public override string ToString()
            {
                return String.Format("{0} - {1} Items", MainClassName, ChildPath.Count);
            }
        }

        internal class RunningWindowTableItemCollection : SortableBindingList<ProxyInformation>, IDisposableEnumeration<ProxyInformation>
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
                        catch (ObjectDisposedException)
                        {
                            // not nice but i didnt find a way to check an IDisposable instance is already disposed 
                            // - SL
                            ;
                        }
                        catch
                        {
                            throw;
                        }

                        if (item is MarshalByRefObject)
                        {
                            Marshal.ReleaseComObject(item);
                        }
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
    }
}
