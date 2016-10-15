using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Encapsulate a Proxy and spend some additional informations about
    /// </summary>
    public class ProxyInformation : IDisposable
    {
        #region Nested

        /// <summary>
        /// Determine Process Elevation
        /// </summary>
        public enum ProcessElevation
        {
            /// <summary>
            /// Failed to detect permission
            /// </summary>
            Unknown = 0,

            /// <summary>
            /// Process is in admin role
            /// </summary>
            AdministratorRole = 1,

            /// <summary>
            /// Process is not in admin role
            /// </summary>
            BelowAdministratorRole = 2
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="proxy">proxy from ROT</param>
        /// <param name="displayName">display name in running object table</param>
        /// <param name="id">interface id</param>
        /// <param name="name">name of the managed proxy class if exists</param>
        /// <param name="component">name of the component where the proxy comes from</param>
        /// <param name="libraryID">id of the component where the proxy comes from</param>
        /// <param name="processID">pid</param>
        /// <param name="elevation">process elevation</param>
        public ProxyInformation(object proxy, string displayName, string id, string name, 
            string component, string libraryID, IntPtr processID, ProcessElevation elevation)
        {
            if (null == proxy)
                throw new ArgumentNullException("proxy");
            Proxy = proxy;
            DisplayName = displayName;
            ID = id == Guid.Empty.ToString() ? "<Unknown>" : id;
            Name = String.IsNullOrWhiteSpace(name) ? "<Unknown>" : name;
            Component = String.IsNullOrWhiteSpace(component) ? "<Unknown>" : component;
            Library = libraryID == Guid.Empty.ToString() ? "<Unknown>" : libraryID;
            ProcessID = processID;
            Elevation = elevation;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Display name in the running object table
        /// </summary>
        [Category("Details")]
        public string DisplayName { get; private set; }

        /// <summary>
        /// Name of the managed proxy class if exists
        /// </summary>
        [Category("Details")]
        public string Name { get; private set; }

        /// <summary>
        /// Name of the component where the proxy comes from
        /// </summary>
        [Category("Details")]
        public string Component { get; private set; }

        /// <summary>
        /// Interface id
        /// </summary>
        [Category("Details")]
        public string ID { get; private set; }

        /// <summary>
        /// ID of the component where the proxy comes from
        /// </summary>
        [Category("Details")]
        public string Library { get; private set; }

        /// <summary>
        /// PID
        /// </summary>
        [Category("Details")]
        public IntPtr ProcessID { get; private set; }

        /// <summary>
        /// Determine process elevation
        /// </summary>
        [Category("Details")]
        public ProcessElevation Elevation { get; private set; }

        /// <summary>
        /// Proxy from ROT
        /// </summary>
        [Browsable(false)]
        public object Proxy { get; private set; }

        #endregion

        #region IDisposable

        /// <summary>
        /// Release the proxy
        /// </summary>
        public void Dispose()
        {
            if (null != Proxy && Proxy is MarshalByRefObject)
            {
                Marshal.ReleaseComObject(Proxy);
                Proxy = null;
            }
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.IsNullOrWhiteSpace(DisplayName) ? base.ToString() : DisplayName;
        }

        #endregion
    }
}
