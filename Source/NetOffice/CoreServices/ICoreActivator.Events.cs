using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices
{
    #region Create Instance

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
            RequestedFrom = instance;
        }

        /// <summary>
        /// The instance candidate to replace.
        /// DisposeChildInstances is called for the instance after event triger
        /// </summary>
        public ICOMObject RequestedFrom { get; private set; }

        /// <summary>
        /// Type muste inherit from origin instance interface and make empty public ctor available
        /// </summary>
        public ICOMObject Replace { get; set; }
    }

    /// <summary>
    /// OnCreateInstance event handler
    /// </summary>
    /// <param name="sender">Core sender instance</param>
    /// <param name="args">args as provided</param>
    public delegate void OnCreateInstanceEventHandler(Core sender, OnCreateInstanceEventArgs args);

    #endregion

    #region Create COMDynamic

    /// <summary>
    /// Arguments in CreateCOMDynamicEvent event
    /// </summary>
    public class OnCreateCOMDynamicEventArgs : EventArgs
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="requestedFrom">calling wrapper instance</param>
        /// <param name="comProxy">target proxy</param>
        internal OnCreateCOMDynamicEventArgs(ICOMObject requestedFrom, object comProxy)
        {
            RequestedFrom = requestedFrom;
            ComProxy = comProxy;
        }

        /// <summary>
        /// Calling wrapper instance
        /// </summary>
        public ICOMObject RequestedFrom { get; private set; }

        /// <summary>
        /// Target Proxy
        /// </summary>
        public object ComProxy { get; private set; }

        /// <summary>
        /// COMDynamicObject instance to set or null for default
        /// </summary>
        public COMDynamicObject Replace { get; set; }
    }

    /// <summary>
    /// OnCreateCOMDynamic event handler
    /// </summary>
    /// <param name="sender">Core sender instance</param>
    /// <param name="args">args as provided</param>
    public delegate void OnCreateCOMDynamicEventHandler(Core sender, OnCreateCOMDynamicEventArgs args);

    #endregion

    #region Create ProxyShare

    /// <summary>
    /// Arguments in CreateProxyShare event
    /// </summary>
    public class OnCreateProxyShareEventArgs : EventArgs
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="requestedFrom">calling wrapper instance</param>
        /// <param name="isEnumerator">indicates rcw is an enumerator</param>
        internal OnCreateProxyShareEventArgs(ICOMObject requestedFrom, bool isEnumerator)
        {
            RequestedFrom = requestedFrom;
            IsEnumerator = IsEnumerator;
        }

        /// <summary>
        /// Calling wrapper instance
        /// </summary>
        public ICOMObject RequestedFrom { get; private set; }

        /// <summary>
        /// Indicates rcw is an enumerator
        /// </summary>
        public bool IsEnumerator { get; private set; }

        /// <summary>
        /// COMProxyShare instance to set or null for default
        /// </summary>
        public COMProxyShare Result { get; set; }
    }

    /// <summary>
    /// OnCreateProxy event handler
    /// </summary>
    /// <param name="sender">Core sender instance</param>
    /// <param name="args">args as provided</param>
    public delegate void OnCreateProxyShareEventHandler(Core sender, OnCreateProxyShareEventArgs args);

    #endregion
}
