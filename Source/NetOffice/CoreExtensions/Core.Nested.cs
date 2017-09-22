using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace NetOffice
{
    /// <summary>
    /// Nested Core Declarations
    /// </summary>
    partial class Core
    {
        #region Nested

        /// <summary>
        /// Provides detailed informations about a com proxy
        /// </summary>
        public class ProxyInformations
        {
            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="name"></param>
            /// <param name="fullComponentName"></param>
            /// <param name="typeID"></param>
            public ProxyInformations(string name, string fullComponentName, Guid typeID)
            {
                Name = name;
                FullComponentName = fullComponentName;
                TypeID = typeID;
            }

            /// <summary>
            /// Class Name
            /// </summary>
            public string Name { get; private set; }

            /// <summary>
            /// Component Name
            /// </summary>
            public string FullComponentName { get; private set; }

            /// <summary>
            /// Type/Class ID
            /// </summary>
            public Guid TypeID { get; private set; }

            /// <summary>
            /// Creates new instance of the class
            /// </summary>
            /// <param name="comProxy">target proxy</param>
            /// <returns>ProxyInformations instance</returns>
            public static ProxyInformations Create(object comProxy)
            {
                string className = TypeDescriptor.GetClassName(comProxy);
                string componentName = TypeDescriptor.GetComponentName(comProxy);
                Guid typeID = comProxy.TypeGuid();
                return new ProxyInformations(className, componentName, typeID);
            }
        }

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
            public COMDynamicObject Result { get; set; }
        }

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
                Instance = instance;
            }

            /// <summary>
            /// The instance candidate to replace.
            /// DisposeChildInstances is called for the instance after event triger
            /// </summary>
            public ICOMObject Instance { get; private set; }

            /// <summary>
            /// Type muste inherit from origin instance class type and make COMObject public .ctors available
            /// </summary>
            public Type Replace { get; set; }
        }

        /// <summary>
        /// Arguments in Resolve Event
        /// </summary>
        public class ResolveEventArgs
        {
            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="caller">calling instance</param>
            /// <param name="fullClassName">target NetOffice class</param>
            /// <param name="comProxy">native proxy type</param>
            public ResolveEventArgs(ICOMObject caller, string fullClassName, Type comProxy)
            {
                Caller = caller;
                FullClassName = fullClassName;
                ComProxy = ComProxy;
            }

            /// <summary>
            /// Calling instance or null(Nothing in Visual Basic)
            /// </summary>
            public ICOMObject Caller { get; private set; }

            /// <summary>
            /// Target NetOffice class as full qualified name
            /// </summary>
            public string FullClassName { get; private set; }

            /// <summary>
            /// Native Proxy Type
            /// </summary>
            public Type ComProxy { get; private set; }

            /// <summary>
            /// Wrapper class to create an instance from 
            /// </summary>
            public Type Result { get; set; }
        }

        /// <summary>
        /// Resolve event handler
        /// </summary>
        /// <param name="sender">Core sender instance</param>
        /// <param name="args">args as provided</param>
        public delegate void ResolveEventHandler(Core sender, ResolveEventArgs args);

        /// <summary>
        /// OnCreateCOMDynamic event handler
        /// </summary>
        /// <param name="sender">Core sender instance</param>
        /// <param name="args">args as provided</param>
        public delegate void OnCreateCOMDynamicEventHandler(Core sender, OnCreateCOMDynamicEventArgs args);

        /// <summary>
        /// OnCreateProxy event handler
        /// </summary>
        /// <param name="sender">Core sender instance</param>
        /// <param name="args">args as provided</param>
        public delegate void OnCreateProxyShareEventHandler(Core sender, OnCreateProxyShareEventArgs args);

        /// <summary>
        /// OnCreateInstance event handler
        /// </summary>
        /// <param name="sender">Core sender instance</param>
        /// <param name="args">args as provided</param>
        public delegate void OnCreateInstanceEventHandler(Core sender, OnCreateInstanceEventArgs args);

        /// <summary>
        /// ProxyCountChanged delegate
        /// </summary>
        /// <param name="proxyCount">current count of com proxies</param>
        public delegate void ProxyCountChangedHandler(int proxyCount);

        /// <summary>
        /// IsInitializedChanged delegate
        /// </summary>
        /// <param name="isInitialized"></param>
        public delegate void IsInitializedChangedHandler(bool isInitialized);

        /// <summary>
        /// Proxy added delegate
        /// </summary>
        /// <param name="sender">sender instance</param>
        /// <param name="ownerPath">comObject relation path</param>
        /// <param name="comObject">added object</param>
        public delegate void ProxyAddedHandler(Core sender, IEnumerable<ICOMObject> ownerPath, ICOMObject comObject);

        /// <summary>
        /// Proxy remove delegate
        /// </summary>
        /// <param name="sender">sender instance</param>
        /// <param name="ownerPath">former comObject relation path</param>
        /// <param name="comObject">removed object</param>
        public delegate void ProxyRemovedHandler(Core sender, IEnumerable<ICOMObject> ownerPath, ICOMObject comObject);

        /// <summary>
        /// Proxy clear delegate
        /// </summary>
        /// <param name="sender">sender instance</param>
        public delegate void ProxyClearHandler(Core sender);
        
        #endregion
    }
}
