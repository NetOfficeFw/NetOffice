using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Collections.Generic;
using System.Text;
using COMTypes = System.Runtime.InteropServices.ComTypes;

namespace NetOffice
{
    /// <summary>
    /// represents a managed COM proxy 
    /// </summary>
    [TypeConverter(typeof(COMObjectExpandableObjectConverter))]
    public class COMObject : IDisposable
    {
        #region Fields

        /// <summary>
        /// the well know IUnknown Interface ID
        /// </summary>
        private static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        /// <summary>
        /// returns parent instance
        /// </summary>
        protected internal COMObject _parentObject;

        /// <summary>
        /// returns Type of native proxy
        /// </summary>
        protected internal Type _instanceType;

        /// <summary>
        /// returns the native wrapped proxy
        /// </summary>
        protected internal object _underlyingObject;

        /// <summary>
        /// returns instance is an enumerator
        /// </summary>
        protected internal bool _isEnumerator;

        /// <summary>
        /// returns instance implement quit method and dispose call them automaticly
        /// </summary>
        protected internal bool _callQuitInDispose;

        /// <summary>
        /// returns instance is currently in disposing progress
        /// </summary>
        protected internal volatile bool _isCurrentlyDisposing;

        /// <summary>
        /// returns instance is diposed means unusable
        /// </summary>
        protected internal volatile bool _isDisposed;

        /// <summary>
        ///  child instance List
        /// </summary>
        protected internal List<COMObject> _listChildObjects = new List<COMObject>();

        /// <summary>
        /// list of runtime supported entities
        /// </summary>
        private Dictionary<string, string> _listSupportedEntities;

        /// <summary>
        /// monitor lock object for accessing the child list
        /// </summary>
        private object _childListLock = new object();

        #endregion

        #region Ctor

        /// <summary>
        /// creates instance and replace the given replacedObject in proxy management
        /// all created childs from replacedObject are now childs from the new instance
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="replacedObject">the instance you want replace in current NO proxy management</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, COMObject replacedObject)
        {
            // copy current factory info or set default
            if (null == factory)
                factory = Core.Default;
            Factory = factory;

            // copy proxy
            _underlyingObject = replacedObject.UnderlyingObject;
            _parentObject = replacedObject.ParentObject;
            _instanceType = replacedObject.InstanceType;

            // copy childs
            foreach (COMObject item in replacedObject.ListChildObjects)
                AddChildObject(item);

            // remove old object from parent chain
            if (!Object.ReferenceEquals(replacedObject.ParentObject, null))
            {
                COMObject parentObject = replacedObject.ParentObject;
                parentObject.RemoveChildObject(replacedObject);

                // add himself as child to parent object
                parentObject.AddChildObject(this);
            }

            Factory.RemoveObjectFromList(replacedObject);
            Factory.AddObjectToList(this);
        }

        /// <summary>
        /// creates instance and replace the given replacedObject in proxy management
        /// all created childs from replacedObject are now childs from the new instance
        /// </summary>
        /// <param name="replacedObject">the instance you want replace in current NO proxy management</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(COMObject replacedObject)
        {
            // copy current factory info or set default
            if (null != replacedObject)
                Factory = replacedObject.Factory;
            else
                Factory = Core.Default;

            // copy proxy
            _underlyingObject = replacedObject.UnderlyingObject;
            _parentObject = replacedObject.ParentObject;
            _instanceType = replacedObject.InstanceType;

            // copy childs
            foreach (COMObject item in replacedObject.ListChildObjects)
                AddChildObject(item);

            // remove old object from parent chain
            if (!Object.ReferenceEquals(replacedObject.ParentObject, null))
            {
                COMObject parentObject = replacedObject.ParentObject;
                parentObject.RemoveChildObject(replacedObject);

                // add himself as child to parent object
                parentObject.AddChildObject(this);
            }

            Factory.RemoveObjectFromList(replacedObject);
            Factory.AddObjectToList(this);
        }

        /// <summary>
        /// creates new instance with given proxy
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="comProxy">the now wrapped comProxy root instance</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, object comProxy)
        {
            // copy current factory info or set default
            if (null == factory)
                factory = Core.Default;
            Factory = factory;

            _underlyingObject = comProxy;
            _instanceType = comProxy.GetType();

            Factory.AddObjectToList(this);
        }

        /// <summary>
        /// creates new instance with given proxy and parent info
        /// </summary>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(COMObject parentObject, object comProxy)
        {
            // copy current factory info or set default
            if (null != parentObject)
                Factory = parentObject.Factory;
            else
                Factory = Core.Default;

            _parentObject = parentObject;
            _underlyingObject = comProxy;
            _instanceType = comProxy.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);
        }

        /// <summary>
        /// creates new instance with given proxy
        /// </summary>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(object comProxy)
        {
            // copy current factory info or set default
            Factory = Core.Default;

            _parentObject = null;
            _underlyingObject = comProxy;
            _instanceType = comProxy.GetType();

            Factory.AddObjectToList(this);
        }

        /// <summary>
        /// creates new instance with given proxy and parent info
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, COMObject parentObject, object comProxy)
        {
            // copy current factory info or set default
            if (null == factory)
                factory = Core.Default;
            Factory = factory;

            _parentObject = parentObject;
            _underlyingObject = comProxy;
            _instanceType = comProxy.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);
        }

        /// <summary>
        /// creates new instance with given proxy, parent info and info instance is an enumerator
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        ///  <param name="isEnumerator"></param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, COMObject parentObject, object comProxy, bool isEnumerator)
        {
            // copy current factory info
            if (null == factory)
                factory = Core.Default;
            Factory = factory;

            _parentObject = parentObject;
            _underlyingObject = comProxy;
            _isEnumerator = isEnumerator;
            _instanceType = comProxy.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);
        }

        /// <summary>
        /// creates new instance with given proxy, type info and parent info
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        /// <param name="comProxyType">typeinfo from comProy if you have or null</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, COMObject parentObject, object comProxy, Type comProxyType)
        {
            // copy current factory info
            if (null == factory)
                factory = Core.Default;
            Factory = factory;

            _parentObject = parentObject;
            _underlyingObject = comProxy;

            if (null != comProxyType)
                _instanceType = comProxyType;
            else
                comProxyType.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);
        }

        /// <summary>
        /// creates new instance with given proxy, type info and parent info
        /// </summary>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        /// <param name="comProxyType">typeinfo from comProy if you have or null</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(COMObject parentObject, object comProxy, Type comProxyType)
        {
            // copy current factory info or set default
            if (null != parentObject)
                Factory = parentObject.Factory;
            else
                Factory = Core.Default;

            _parentObject = parentObject;
            _underlyingObject = comProxy;

            if (null != comProxyType)
                _instanceType = comProxyType;
            else
                comProxyType.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);
        }

        /// <summary>
        /// creates a new instace with progid
        /// </summary>
        /// <param name="factory">current factory instance</param>
        /// <param name="progId">registered ProgID</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, string progId)
        {
            // copy current factory info
            if (null == factory)
                factory = Core.Default;
            Factory = factory;

            CreateFromProgId(progId);
            Factory.AddObjectToList(this);
        }

        /// <summary>
        /// creates a new instace with progid
        /// </summary>
        /// <param name="progId">registered ProgID</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(string progId)
        {
            CreateFromProgId(progId);
            Factory = Core.Default;
            Factory.AddObjectToList(this);
        }

        /// <summary>
        /// not usable stub constructor
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public COMObject()
        {
            DebugConsole.Default.WriteLine("Warning: Invalid COMObject Stub Ctor called.");
        }

        #endregion

        #region COMObject Properties

        /// <summary>
        /// NetOffice property: the associated factory
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced), Category("NetOffice")]
        public Core Factory
        {
            get
            {
                if (null == _factory)
                    return Core.Default;
                else
                    return _factory;
            }
            set
            {
                _factory = value;
            }
        }
        private Core _factory;

        /// <summary>
        /// NetOffice property: the associated invoker
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced), Category("NetOffice")]
        public Invoker Invoker
        {
            get
            {
                if (null != _factory)
                    return _factory.Invoker;
                else
                    return Invoker.Default;
            }
        }

        /// <summary>
        /// NetOffice property: the associated console
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced), Category("NetOffice")]
        public DebugConsole Console
        {
            get
            {
                if (null != _factory)
                    return _factory.Console;
                else
                    return DebugConsole.Default;
            }
        }

        /// <summary>
        /// NetOffice property: the associated settings
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced), Category("NetOffice")]
        public Settings Settings
        {
            get
            {
                if (null != _factory)
                    return _factory.Settings;
                else
                    return Settings.Default;
            }
        }

        /// <summary>
        /// NetOffice property: returns the native wrapped proxy
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced), Category("NetOffice")]
        public object UnderlyingObject
        {
            get
            {
                return _underlyingObject;
            }
        }

        /// <summary>
        /// NetOffice property: returns friendly name for the instance type
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced), Category("NetOffice")]
        public string FriendlyTypeName
        {
            get
            {
                string fullname = InstanceType.FullName;
                if (!String.IsNullOrEmpty(fullname))
                {
                    if (fullname.StartsWith("NetOffice.", StringComparison.InvariantCultureIgnoreCase))
                        fullname = fullname.Replace("NetOffice.", "");
                    if (fullname.StartsWith("Api.", StringComparison.InvariantCultureIgnoreCase))
                        fullname = fullname.Replace("Api.", "");
                }
                return fullname;
            }
        }

        /// <summary>
        /// NetOffice property: returns class name of native wrapped proxy
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false), Category("NetOffice")]
        public string UnderlyingTypeName
        {
            get
            {
                return TypeDescriptor.GetClassName(_underlyingObject);
            }
        }

        /// <summary>
        /// NetOffice property: returns component name of native wrapped proxy
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false), Category("NetOffice")]
        public string UnderlyingComponentName
        {
            get
            {
                return TypeDescriptor.GetComponentName(_underlyingObject);
            }
        }

        /// <summary>
        /// NetOffice property: returns instance is diposed means unusable
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        public bool IsDisposed
        {
            get
            {
                return _isDisposed;
            }
        }

        /// <summary>
        /// NetOffice property: returns Type of native proxy
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false), Category("NetOffice")]
        public Type InstanceType
        {
            get
            {
                return _instanceType;
            }
        }

        /// <summary>
        /// NetOffice property: returns parent proxy object
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        public COMObject ParentObject
        {
            get
            {
                return _parentObject;
            }
            set
            {
                _parentObject = value;
            }
        }
        /// <summary>
        /// NetOffice property: returns instance is currently in diposing progress
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        public bool IsCurrentlyDisposing
        {
            get
            {
                return _isCurrentlyDisposing;
            }
        }

        /// <summary>
        ///  child instance array
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false), Category("NetOffice")]
        internal COMObject[] ListChildObjects
        {
            get
            {
                return _listChildObjects.ToArray();
            }
        }

        /// <summary>
        /// NetOffice property: returns instance export events
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        public bool IsEventBinding
        {
            get
            {
                return (!Object.ReferenceEquals(this as IEventBinding, null));
            }
        }

        /// <summary>
        /// NetOffice property: returns event bridge is advised
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        public bool IsEventBridgeInitialized
        {
            get
            {
                IEventBinding bindInfo = this as IEventBinding;
                if (!Object.ReferenceEquals(bindInfo, null))
                    return bindInfo.EventBridgeInitialized;
                else
                    return false;
            }
        }

        /// <summary>
        /// NetOffice property: retuns instance has one or more event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        public bool IsWithEventRecipients
        {
            get
            {
                IEventBinding bindInfo = this as IEventBinding;
                if (!Object.ReferenceEquals(bindInfo, null))
                    return bindInfo.HasEventRecipients();
                else
                    return false;
            }
        }

        #endregion

        #region COMObject Methods

        /// <summary>
        /// NetOffice method: returns information the proxy provides a method or property with given name at runtime
        /// </summary>
        /// <param name="name">name of the enitity</param>
        /// <returns>true if available, otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool EntityIsAvailable(string name)
        {
            return EntityIsAvailable(name, SupportEntityType.Both);
        }

        /// <summary>
        ///  NetOffice method: returns information the proxy provides a method or property with given name at runtime
        /// </summary>
        /// <param name="name">name of the enitity</param>
        /// <param name="searchType">limit searching for method or property</param>
        /// <returns>true if available, otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool EntityIsAvailable(string name, SupportEntityType searchType)
        {
            switch (searchType)
            {
                case SupportEntityType.Method:
                {
                    if (null == _listSupportedEntities)
                        _listSupportedEntities = Factory.GetSupportedEntities(_underlyingObject);

                    string outValue = null;
                    return _listSupportedEntities.TryGetValue("Method-" + name, out outValue);   
                }
                case SupportEntityType.Property:
                {
                    if (null == _listSupportedEntities)
                        _listSupportedEntities = Factory.GetSupportedEntities(_underlyingObject);

                    string outValue = null;
                    return _listSupportedEntities.TryGetValue("Property-" + name, out outValue);       
                }
                default:
                {
                    if (null == _listSupportedEntities)
                        _listSupportedEntities = Factory.GetSupportedEntities(_underlyingObject);

                    string outValue = null;
                    bool result = _listSupportedEntities.TryGetValue("Property-" + name, out outValue);
                    if (result)
                        return true;

                    return _listSupportedEntities.TryGetValue("Method-" + name, out outValue);    
                }
            }
        }

        /// <summary>
        /// NetOffice method: create object from progid
        /// </summary>
        /// <param name="progId"></param>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public void CreateFromProgId(string progId)
        {
            _instanceType = System.Type.GetTypeFromProgID(progId);
            if (null == _instanceType)
                throw (new ArgumentException("progId not found. " + progId));

            _underlyingObject = Activator.CreateInstance(_instanceType);
        }

        /// <summary>
        ///  NetOffice method: add object to child list
        /// </summary>
        /// <param name="childObject">target child instance</param>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public void AddChildObject(COMObject childObject)
        {
            bool isLocked = false;
            try
            {
                Monitor.Enter(_childListLock);
                isLocked = true;

                _listChildObjects.Add(childObject);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw (throwedException);
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_childListLock);
                    isLocked = false;
                }
            }
        }

        /// <summary>
        /// remove object to child list
        /// </summary>
        /// <param name="childObject">target child instance</param>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public void RemoveChildObject(COMObject childObject)
        {
            bool isLocked = false;
            try
            {
                _listChildObjects.Remove(childObject);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw (throwedException);
            }
            finally
            {
                if (isLocked)
                {
                    Monitor.Exit(_childListLock);
                    isLocked = false;
                }
            }
        }

        /// <summary>
        ///  NetOffice method: release com proxy
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public void ReleaseCOMProxy()
        {
            // release himself from COM Runtime System
            if (!Object.ReferenceEquals(_underlyingObject, null))
            {
                if (_isEnumerator)
                {
                    ICustomAdapter adapter = _underlyingObject as ICustomAdapter;
                    Marshal.ReleaseComObject(adapter.GetUnderlyingObject());
                }
                else
                {
                    Marshal.ReleaseComObject(_underlyingObject);
                }
                Factory.RemoveObjectFromList(this);
                _underlyingObject = null;
            }
        }

        /// <summary>
        /// calls Quit for a proxy
        /// </summary>
        /// <param name="settings"></param>
        /// <param name="invoker"></param>
        /// <param name="instance"></param>
        private static void CallQuit(Settings settings, Invoker invoker, COMObject instance)
        {
            try
            {
                if (null == instance)
                    return;
                if (settings.EnableAutomaticQuit)
                    invoker.Method(instance.UnderlyingObject, "Quit");
            }
            catch (Exception exception)
            {
                instance.Console.WriteException(exception);
            }
        }

        #endregion

        #region IDisposable Members

        /// <summary>
        /// NetOffice event: these event was called from Dispose and you can skip the dipose operation here if you want. the event can be helpful for troubleshooting if you dont know why your objects beeing disposed
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public event OnDisposeEventHandler OnDispose;

        /// <summary>
        /// Calls the OnDispose event as service for client callers
        /// </summary>
        /// <returns></returns>
        private bool RaiseOnDispose()
        {
            bool cancelDispose = false;
            try
            {
                if (null != OnDispose)
                {
                    OnDisposeEventArgs eventArgs = new OnDisposeEventArgs(this);
                    OnDispose(eventArgs);
                    cancelDispose = eventArgs.Cancel;
                }
            }
            catch (Exception exception)
            {
                Console.WriteException(exception);
            }
            return cancelDispose;
        }

        /// <summary>
        /// NetOffice method: dispose instance and all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose event exported proxies with one or more event recipients</param>
        public virtual void Dispose(bool disposeEventBinding)
        {
            // skip check 
            bool cancel = RaiseOnDispose();
            if (cancel)
                return;

            // in case object export events and 
            // disposeEventBinding == true we dont remove the object from parents child list
            bool removeFromParent = true;

            // set disposed flag
            _isCurrentlyDisposing = true;

            // in case of object implements also event binding we dispose them
            IEventBinding eventBind = this as IEventBinding;
            if (disposeEventBinding)
            {
                if (!Object.ReferenceEquals(eventBind, null))
                    eventBind.DisposeEventBridge();
            }
            else
            {
                if (!Object.ReferenceEquals(eventBind, null) && (eventBind.EventBridgeInitialized))
                    removeFromParent = false;
            }

            // child proxy dispose
            DisposeChildInstances(disposeEventBinding);


            // remove himself from parent childlist
            if ((!Object.ReferenceEquals(_parentObject, null)) && (true == removeFromParent))
            {
                _parentObject.RemoveChildObject(this);
                _parentObject = null;
            }

            // call quit automaticly if wanted
            if (_callQuitInDispose)
                CallQuit(Settings, Invoker, this);

            // release proxy
            ReleaseCOMProxy();

            // clear supportList reference
            _listSupportedEntities = null;

            _isDisposed = true;
            _isCurrentlyDisposing = false;
        }

        /// <summary>
        /// NetOffice method: dispose instance and all child instances
        /// </summary>
        public virtual void Dispose()
        {
            Dispose(true);
        }

        /// <summary>
        /// NetOffice method: dispose all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose proxies with events and one or more event recipients</param>
        public void DisposeChildInstances(bool disposeEventBinding)
        {
            // release all childs and clear list
            foreach (COMObject itemObject in _listChildObjects)
            {
                itemObject.ParentObject = null;
                itemObject.Dispose(disposeEventBinding);
            }
            _listChildObjects.Clear();
        }

        /// <summary>
        /// NetOffice method: dispose all child instances
        /// </summary>
        public void DisposeChildInstances()
        {
            // release all childs and clear list
            foreach (COMObject itemObject in _listChildObjects)
            {
                itemObject.ParentObject = null;
                itemObject.Dispose();
            }
            _listChildObjects.Clear();
        }

        #endregion

        #region Object overrides

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>System.Int32 instance</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        /// <summary>
        /// Returns a string that represents the current object.
        /// </summary>
        /// <returns>System.String instance</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public override string ToString()
        {
            return GetType().Name;
        }

        /// <summary>
        /// Determines whether two Object instances are equal.
        /// </summary>
        /// <returns>true if equal, otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public override bool Equals(Object obj)
        {
            return base.Equals(obj);
        }

        /// <summary>
        /// Determines whether two COMObject instances are equal.
        /// </summary>
        /// <param name="obj">target instance to compare</param>
        /// <returns>true if equal, otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool EqualsOnServer(COMObject obj)
        {
            if (_isCurrentlyDisposing || _isDisposed)
                return base.Equals(obj);

            if (Object.ReferenceEquals(obj, null))
                return false;

            IntPtr outValueA = IntPtr.Zero;
            IntPtr outValueB = IntPtr.Zero;
            IntPtr ptrA = IntPtr.Zero;
            IntPtr ptrB = IntPtr.Zero;
            try
            {
                ptrA = Marshal.GetIUnknownForObject(this.UnderlyingObject);
                int hResultA = Marshal.QueryInterface(ptrA, ref IID_IUnknown, out outValueA);

                ptrB = Marshal.GetIUnknownForObject(obj.UnderlyingObject);
                int hResultB = Marshal.QueryInterface(ptrB, ref IID_IUnknown, out outValueB);

                return (hResultA == 0 && hResultB == 0 && ptrA == ptrB);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
                throw exception;
            }
            finally
            {
                if (IntPtr.Zero != ptrA)
                    Marshal.Release(ptrA);

                if (IntPtr.Zero != outValueA)
                    Marshal.Release(outValueA);

                if (IntPtr.Zero != ptrB)
                    Marshal.Release(ptrB);

                if (IntPtr.Zero != outValueB)
                    Marshal.Release(outValueB);
            }
        }

        /// <summary>
        /// Gets a Type object that represents the specified type.
        /// </summary>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public new Type GetType()
        {
            return base.GetType();
        }

        #endregion

        #region Operator Overloads

        /// <summary>
        /// Determines whether two COMObject instances are equal.
        /// </summary>
        /// <param name="objectA"></param>
        /// <param name="objectB"></param>
        /// <returns></returns>
        public static bool operator ==(COMObject objectA, COMObject objectB)
        {
            if (!Settings.Default.EnableOperatorOverlads)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return true;
            else if (!Object.ReferenceEquals(objectA, null))
                return objectA.EqualsOnServer(objectB);
            else
                return false;
        }

        /// <summary>
        /// Determines whether two COMObject instances are equal.
        /// </summary>
        /// <param name="objectA"></param>
        /// <param name="objectB"></param>
        /// <returns></returns>
        public static bool operator ==(COMObject objectA, object objectB)
        {
            if (!Settings.Default.EnableOperatorOverlads)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return true;
            else if (!Object.ReferenceEquals(objectA, null))
                return objectA.EqualsOnServer(objectB as COMObject);
            else
                return false;
        }

        /// <summary>
        /// Determines whether two COMObject instances are equal.
        /// </summary>
        /// <param name="objectA">first instance</param>
        /// <param name="objectB">second instance</param>
        /// <returns>true if equal, otherwise false</returns>
        public static bool operator ==(object objectA, COMObject objectB)
        {
            if (!Settings.Default.EnableOperatorOverlads)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return true;
            else if (!Object.ReferenceEquals(objectA, null))
                return (objectA as COMObject).EqualsOnServer(objectB);
            else
                return false;
        }

        /// <summary>
        /// Determines whether two COMObject instances are not equal.
        /// </summary>
        /// <param name="objectA">first instance</param>
        /// <param name="objectB">second instance</param>
        /// <returns>true if equal, otherwise false</returns>
        public static bool operator !=(COMObject objectA, COMObject objectB)
        {
            if (!Settings.Default.EnableOperatorOverlads)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return false;
            else if (!Object.ReferenceEquals(objectA, null))
                return !objectA.EqualsOnServer(objectB);
            else
                return true;
        }

        /// <summary>
        /// Determines whether two COMObject instances are not equal.
        /// </summary>
        /// <param name="objectA">first instance</param>
        /// <param name="objectB">second instance</param>
        /// <returns>true if equal, otherwise false</returns>
        public static bool operator !=(COMObject objectA, object objectB)
        {
            if (!Settings.Default.EnableOperatorOverlads)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return false;
            else if (!Object.ReferenceEquals(objectA, null))
                return !objectA.EqualsOnServer(objectB as COMObject);
            else
                return true;
        }

        /// <summary>
        /// Determines whether two COMObject instances are not equal.
        /// </summary>
        /// <param name="objectA">first instance</param>
        /// <param name="objectB">second instance</param>
        /// <returns>true if equal, otherwise false</returns>
        public static bool operator !=(object objectA, COMObject objectB)
        {
            if (!Settings.Default.EnableOperatorOverlads)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return false;
            else if (!Object.ReferenceEquals(objectA, null))
                return !(objectA as COMObject).EqualsOnServer(objectB);
            else
                return true;
        }

        #endregion
    }
}