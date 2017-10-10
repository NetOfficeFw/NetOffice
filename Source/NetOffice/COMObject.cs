using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Collections.Generic;
using NetOffice.Resolver;
using NetOffice.Availity;
using NetOffice.Exceptions;

namespace NetOffice
{
    /// <summary>
    /// Managed/wrapped COM Proxy and ICOMObject default implementation
    /// </summary>
    [DebuggerDisplay("{InstanceFriendlyName}")]
    [TypeConverter(typeof(Converter.COMObjectExpandableObjectConverter))]
    public class COMObject : ICOMObject, ICOMProxyShareProvider
    {
        #region Fields

        /// <summary>
        /// Create from ProgId Failed Message
        /// </summary>
        private static readonly string _createFromProgIdFailMessageHint = "This is typically because you have no access to the desktop subsystem " +
                                                                   "from a Windows Service/IIS modul in default configuration because its running in a restricted context/principal.";

        /// <summary>
        /// the well know IUnknown Interface ID
        /// </summary>
        private static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        /// <summary>
        /// returns parent instance
        /// </summary>
        protected internal ICOMObject _parentObject;

        /// <summary>
        /// returns Type of native proxy
        /// </summary>
        protected internal Type _underlyingType;

        /// <summary>
        /// Returns a shared access wrapper arround the native wrapped RCW System._ComObject
        /// </summary>
        protected internal COMProxyShare _proxyShare;

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
        /// child instance List
        /// </summary>
        protected internal List<ICOMObject> _listChildObjects = new List<ICOMObject>();

        /// <summary>
        /// list of runtime supported entities
        /// </summary>
        private Dictionary<string, string> _listSupportedEntities;
        
        /// <summary>
        /// monitor lock object for the main dispose method
        /// </summary>
        private object _disposeLock = new object();
       
        /// <summary>
        /// monitor lock object for accessing the child list
        /// </summary>
        private object _childListLock = new object();

        /// <summary>
        /// monitor lock object for dispose the child list
        /// </summary>
        private object _disposeChildLock = new object();

        /// <summary>
        /// associated factory
        /// </summary>
        private Core _factory;

        /// <summary>
        /// FriendlyTypeName chache field
        /// </summary>
        private string _friendlyTypeName;

        /// <summary>
        /// UnderlyingTypeName chache field
        /// </summary>
        private string _underlyingTypeName;

        /// <summary>
        /// UnderlyingComponentName chache field
        /// </summary>
        private string _underlyingComponentName;

        /// <summary>
        /// ComponentRootName chache field
        /// </summary>
        private string _componentRootName;

        /// <summary>
        /// InstanceName chache field
        /// </summary>
        private string _instanceName;

        /// <summary>
        /// InstanceType chache field
        /// </summary>
        private Type _instanceType;
        
        #endregion

        #region Ctor

        /// <summary>
        /// Creates instance and replace the given replacedObject in proxy management
        /// all created childs from replacedObject are now childs from the new instance
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="replacedObject">the instance you want replace in current NO proxy management</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, ICOMObject replacedObject)
        {
            if (null == factory)
                factory = Core.Default;
            Factory = factory;
            SyncRoot = new object();

            // copy proxy
            ICOMProxyShareProvider shareProvider = replacedObject as ICOMProxyShareProvider;
            if (null != shareProvider)
                _proxyShare = shareProvider.GetProxyShare();
            else
                _proxyShare = Factory.CreateNewProxyShare(this, replacedObject.UnderlyingObject);
            _parentObject = replacedObject.ParentObject;
            _underlyingType = replacedObject.UnderlyingType;

            // copy childs 
            foreach (ICOMObject item in replacedObject.ChildObjects)
                AddChildObject(item);

            // remove old object from parent chain
            if (!Object.ReferenceEquals(replacedObject.ParentObject, null))
            {
                ICOMObject parentObject = replacedObject.ParentObject;
                parentObject.RemoveChildObject(replacedObject);

                // add himself as child to parent object
                parentObject.AddChildObject(this);
            }

            Factory.RemoveObjectFromList(replacedObject, null);
            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates instance and replace the given replacedObject in proxy management
        /// all created childs from replacedObject are now childs from the new instance
        /// </summary>
        /// <param name="replacedObject">the instance you want replace in current NO proxy management</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(ICOMObject replacedObject)
        {
            if (null != replacedObject)
                Factory = replacedObject.Factory;
            else
                Factory = Core.Default;
            SyncRoot = new object();

            // copy proxy
            ICOMProxyShareProvider shareProvider = replacedObject as ICOMProxyShareProvider;
            if (null != shareProvider)
                _proxyShare = shareProvider.GetProxyShare();
            else
                _proxyShare = Factory.CreateNewProxyShare(this, replacedObject.UnderlyingObject);
            _parentObject = replacedObject.ParentObject;
            _underlyingType = replacedObject.UnderlyingType;

            // copy childs
            foreach (COMObject item in replacedObject.ChildObjects)
                AddChildObject(item);

            // remove old object from parent chain
            if (!Object.ReferenceEquals(replacedObject.ParentObject, null))
            {
                ICOMObject parentObject = replacedObject.ParentObject;
                parentObject.RemoveChildObject(replacedObject);

                // add himself as child to parent object
                parentObject.AddChildObject(this);
            }

            Factory.RemoveObjectFromList(replacedObject, null);
            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates new instance with given proxy
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="comProxy">the now wrapped comProxy root instance</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, object comProxy)
        {         
            if (!(comProxy is MarshalByRefObject))
                throw new ArgumentException("Argument is not a COM proxy." + (null != comProxy ? "(" + comProxy.ToString() + ")" : ""));

            if (null == factory)
                factory = Core.Default;
            Factory = factory;
            SyncRoot = new object();

            _proxyShare = Factory.CreateNewProxyShare(this, comProxy);
            _underlyingType = comProxy.GetType();

            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates new instance with given proxy and parent info
        /// </summary>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(ICOMObject parentObject, object comProxy)
        {
            if (!(comProxy is MarshalByRefObject))
                throw new ArgumentException("Argument is not a COM proxy." + (null != comProxy ? "(" + comProxy.ToString() + ")" : ""));

            if (null != parentObject)
                Factory = parentObject.Factory;
            else
                Factory = Core.Default;
            SyncRoot = new object();

            _parentObject = parentObject;
            _proxyShare = Factory.CreateNewProxyShare(this, comProxy);
            _underlyingType = comProxy.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates new instance with given proxy
        /// </summary>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(object comProxy)
        {
            if (!(comProxy is MarshalByRefObject))
                throw new ArgumentException("Argument is not a COM proxy." + (null != comProxy ? "(" + comProxy.ToString() + ")" : ""));

            Factory = Core.Default;
            SyncRoot = new object();

            _parentObject = null;
            _proxyShare = Factory.CreateNewProxyShare(this, comProxy);
            _underlyingType = comProxy.GetType();

            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates new instance with given proxy and parent info
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="proxyShare">proxy share instead of proxy</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, ICOMObject parentObject, COMProxyShare proxyShare)
        {
            if (null == proxyShare)
                throw new ArgumentNullException("proxyShare");

            if (null == factory)
                factory = Core.Default;
            Factory = factory;
            SyncRoot = new object();

            _parentObject = parentObject;
            _proxyShare = proxyShare;
            _underlyingType = _proxyShare.Proxy.GetType();

            _proxyShare.Acquire();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates new instance with given proxy and parent info
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, ICOMObject parentObject, object comProxy)
        {
            if (!(comProxy is MarshalByRefObject))
                throw new ArgumentException("Argument is not a COM proxy." + (null != comProxy ? "(" + comProxy.ToString() + ")" : ""));

            if (null == factory)
                factory = Core.Default;
            Factory = factory;
            SyncRoot = new object();

            _parentObject = parentObject;
            _proxyShare = Factory.CreateNewProxyShare(this, comProxy);
            _underlyingType = comProxy.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates new instance with given proxy, parent info and info instance is an enumerator
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        ///  <param name="isEnumerator"></param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, ICOMObject parentObject, object comProxy, bool isEnumerator)
        {
            if(false == isEnumerator && (!(comProxy is MarshalByRefObject)))
                throw new ArgumentException("Argument is not a COM proxy." + (null != comProxy ? "(" + comProxy.ToString() + ")" : ""));

            if (null == factory)
                factory = Core.Default;
            Factory = factory;
            SyncRoot = new object();

            _parentObject = parentObject;
            _proxyShare = Factory.CreateNewProxyShare(this, comProxy, isEnumerator);
            _isEnumerator = isEnumerator;
            _underlyingType = comProxy.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates new instance with given proxy, parent info and info instance is an enumerator
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        /// <param name="isEnumerator">instance is an enumerator</param>
        /// <param name="name">custom instance name</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, ICOMObject parentObject, object comProxy, bool isEnumerator, string name)
        {        
            if(false == isEnumerator && (!(comProxy is MarshalByRefObject)))
                throw new ArgumentException("Argument is not a COM proxy." + (null != comProxy ? "(" + comProxy.ToString() + ")" : ""));

            if (null == factory)
                factory = Core.Default;
            Factory = factory;
            SyncRoot = new object();

            _parentObject = parentObject;
            _proxyShare = Factory.CreateNewProxyShare(this, comProxy, isEnumerator);
            _isEnumerator = isEnumerator;
            _underlyingType = comProxy.GetType();
            _instanceName = name;

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates new instance with given proxy, type info and parent info
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        /// <param name="comProxyType">typeinfo from comProy if you have or null</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, ICOMObject parentObject, object comProxy, Type comProxyType)
        {
            if (!(comProxy is MarshalByRefObject))
                throw new ArgumentException("Argument is not a COM proxy." + (null != comProxy ? "(" + comProxy.ToString() + ")" : ""));

            if (null == factory)
                factory = Core.Default;
            Factory = factory;
            SyncRoot = new object();

            _parentObject = parentObject;
            _proxyShare = Factory.CreateNewProxyShare(this, comProxy);

            if (null != comProxyType)
                _underlyingType = comProxyType;
            else
                _underlyingType = comProxy.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates new instance with given proxy, type info and parent info
        /// </summary>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        /// <param name="comProxyType">typeinfo from comProy if you have or null</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(ICOMObject parentObject, object comProxy, Type comProxyType)
        {
            if (!(comProxy is MarshalByRefObject))
                throw new ArgumentException("Argument is not a COM proxy." + (null != comProxy ? "(" + comProxy.ToString() + ")" : ""));

            if (null != parentObject)
                Factory = parentObject.Factory;
            else
                Factory = Core.Default;
            SyncRoot = new object();

            _parentObject = parentObject;
            _proxyShare = Factory.CreateNewProxyShare(this, comProxy);

            if (null != comProxyType)
                _underlyingType = comProxyType;
            else
                _underlyingType = comProxy.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates a new instace with progid
        /// </summary>
        /// <param name="factory">current factory instance</param>
        /// <param name="progId">registered ProgID</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(Core factory, string progId)
        {
            if (String.IsNullOrEmpty(progId))
                throw new ArgumentNullException("progId");

            if (null == factory)
                factory = Core.Default;
            Factory = factory;
            SyncRoot = new object();

            CreateFromProgId(progId);
            Factory.AddObjectToList(this);

            OnCreate();
        }

        /// <summary>
        /// Creates a new instace with progid
        /// </summary>
        /// <param name="progId">registered ProgID</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public COMObject(string progId)
        {
            if (String.IsNullOrEmpty(progId))
                throw new ArgumentNullException("progId");          
            Factory = Core.Default;
            SyncRoot = new object();
            CreateFromProgId(progId);
            Factory.AddObjectToList(this);
             
            OnCreate();
        }

        /// <summary>
        /// Stub
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public COMObject()
        {
            
        }

        #endregion

        #region Properties

        /// <summary>
        /// Always null (Nothing in Visual Basic)
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced), Category("NetOffice")]
        public static COMObject Empty
        {
            get 
            {
                return null;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Creates instance from proxy
        /// </summary>
        /// <typeparam name="T">result type</typeparam>
        /// <param name="comProxy">given proxy as any</param>
        /// <returns>new instance of T</returns>
        /// <exception cref="ArgumentNullException">argument is null(Nothing in Visual Basic)</exception>
        /// <exception cref="ArgumentException">given argument is not a proxy</exception>
        /// <exception cref="CreateInstanceException">unexpected error</exception>
        public static T Create<T>(object comProxy) where T : class, ICOMObject
        {
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");
            if (!(comProxy is MarshalByRefObject))
                throw new ArgumentException("Given argument is not a proxy.");
            try
            {
                return Activator.CreateInstance(typeof(T), new object[] { null, comProxy }) as T;
            }
            catch (Exception exception)
            {
                throw new CreateInstanceException(exception);
            }
        }

        /// <summary>
        /// Creates instance from proxy
        /// </summary>
        /// <typeparam name="T">result type</typeparam>
        /// <param name="factory"></param>
        /// <param name="comProxy">given proxy as any</param>
        /// <returns>new instance of T</returns>
        /// <exception cref="ArgumentNullException">argument is null(Nothing in Visual Basic)</exception>
        /// <exception cref="ArgumentException">given argument is not a proxy</exception>
        /// <exception cref="CreateInstanceException">unexpected error</exception>
        public static T Create<T>(Core factory, object comProxy) where T : class, ICOMObject
        {
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");
            if (!(comProxy is MarshalByRefObject))
                throw new ArgumentException("Given argument is not a proxy.");
            try
            {
                return Activator.CreateInstance(typeof(T), new object[] { factory, null, comProxy }) as T;
            }
            catch (Exception exception)
            {
                throw new CreateInstanceException(exception);
            }   
        }

        /// <summary>
        /// NetOffice method: create object from proxy
        /// </summary>
        /// <param name="underlyingObject">given proxy as any</param>
        /// <param name="factoryAddObject">add instance to factory</param>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        protected void CreateFromProxy(object underlyingObject, bool factoryAddObject = false)
        {
            _underlyingType = underlyingObject.GetType();
            _proxyShare = Factory.CreateNewProxyShare(this, underlyingObject);
            if (factoryAddObject)
                Factory.AddObjectToList(this);
        }

        /// <summary>
        /// NetOffice method: create object from progid
        /// </summary>
        /// <param name="progId">registered progid</param>
        /// <param name="factoryAddObject">add instance to factory</param>
        /// <exception cref="COMException">throws when its failed to resolve progId</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        protected void CreateFromProgId(string progId, bool factoryAddObject = false)
        {
            _underlyingType = System.Type.GetTypeFromProgID(progId, false);
            if (null == _underlyingType)
                throw new COMException("Unable to find registered progId:<" + progId + ">" + Environment.NewLine + _createFromProgIdFailMessageHint);

            object underlyingObject = null;
            try
            {
                underlyingObject = Activator.CreateInstance(_underlyingType);
            }
            catch (Exception exception)
            {
                throw new COMException(
                    "Unable to create instance of:<" + progId + ">" + Environment.NewLine + _createFromProgIdFailMessageHint
                    , exception);
            }

            _proxyShare = Factory.CreateNewProxyShare(this, underlyingObject);
            if (factoryAddObject)
                Factory.AddObjectToList(this);
        }
        
        /// <summary>
        ///  NetOffice method: release com proxy
        /// </summary>
        private void ReleaseCOMProxy(IEnumerable<ICOMObject> ownerPath)
        {
            // release himself from COM Runtime System
            if (!_proxyShare.Released)
            {
                _proxyShare.Release();
                Factory.RemoveObjectFromList(this, ownerPath);
            }
        }

        /// <summary>
        /// Called from ctor at last as an inherited class service
        /// </summary>
        protected internal virtual void OnCreate()
        {

        }

        #endregion

        #region ICOMObject

        /// <summary>
        /// Monitor Lock
        /// </summary>
        public object SyncRoot { get; private set; }

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
        /// Clone instance as target type
        /// </summary>
        /// <typeparam name="T">any other type to convert</typeparam>
        /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public T To<T>() where T : class, ICOMObject
        {
            try
            {
                ICOMObject clone = Activator.CreateInstance(typeof(T), new object[] { Factory, ParentObject, UnderlyingObject }) as ICOMObject;
                ICOMProxyShareProvider shareProvider = clone as ICOMProxyShareProvider;
                shareProvider.SetProxyShare(_proxyShare);
                return clone as T;
            }
            catch (Exception exception)
            {
                throw new CloneException(exception);
            }
        }

        #endregion

        #region ICOMObjectProxy

        /// <summary>
        /// NetOffice property: Returns the native wrapped proxy
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced), Category("NetOffice")]
        public object UnderlyingObject
        {
            get
            {
                return _proxyShare.Proxy;
            }
        }

        /// <summary>
        /// NetOffice property: Type informations from UnderlyingObject
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false), Category("NetOffice")]
        public Type UnderlyingType
        {
            get
            {
                return _underlyingType;
            }
        }

        /// <summary>
        /// NetOffice property: Full type name from UnderlyingObject
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false), Category("NetOffice")]
        public string UnderlyingTypeName
        {
            get
            {
                if (null == _underlyingTypeName)
                    _underlyingTypeName = new UnderlyingTypeNameResolver().GetClassName(this);
                return _underlyingTypeName;
            }
        }

        /// <summary>
        /// NetOffice property: Friendly type name from UnderlyingObject
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced), Category("NetOffice")]
        public string UnderlyingFriendlyTypeName
        {
            get
            {
                if (null == _friendlyTypeName)
                    _friendlyTypeName = new UnderlyingTypeNameResolver().GetFriendlyClassName(this, _underlyingTypeName);
                return _friendlyTypeName;
            }
        }

        /// <summary>
        /// NetOffice property: Component name from UnderlyingObject
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false), Category("NetOffice")]
        public string UnderlyingComponentName
        {
            get
            {
                if (null == _underlyingComponentName)
                    _underlyingComponentName = new UnderlyingTypeNameResolver().GetComponentName(this);
                return _underlyingComponentName;
            }
        }

        /// <summary>
        /// NetOffice property: Full Name of the NetOffice Wrapper class
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false), Category("NetOffice")]
        public string InstanceName
        {
            get
            {
                return InstanceType.FullName;
            }
        }
        /// <summary>
        /// NetOffice property: Friendly Name of the NetOffice Wrapper class
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false), Category("NetOffice")]
        public string InstanceFriendlyName
        {
            get
            {
                if (null == _instanceName)
                    _instanceName = new InstanceTypeNameResolver().GetFriendlyInstanceName(this);
                return _instanceName;
            }
        }

        /// <summary>
        /// NetOffice property: Name of the hosting NetOffice component
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false), Category("NetOffice")]
        public string InstanceComponentName
        {
            get
            {
                if (null == _componentRootName)
                    _componentRootName = new InstanceTypeNameResolver().GetComponentName(this);
                return _componentRootName;
            }
        }

        /// <summary>
        /// NetOffice property: Current Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        public virtual Type InstanceType
        {
            get
            {
                if (null == _instanceType)
                    _instanceType = GetType();
                return _instanceType;
            }
        }

        #endregion

        #region ICOMObjectDisposable
       
        /// <summary>
        /// NetOffice event: these event was called from Dispose and you can skip the dipose operation here if you want. the event can be helpful for troubleshooting if you dont know why your objects beeing disposed
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public event OnDisposeEventHandler OnDispose;
        
        /// <summary>
        /// NetOffice property: returns informations the instance is already disposed
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
        /// NetOffice property: returns information the instance is currently in dispose operation
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
        /// NetOffice method: dispose instance and all child instances
        /// </summary>
        /// <exception cref="COMDisposeException">An unexpected error occurs.</exception>
        public virtual void Dispose()
        {
            Dispose(true);
        }

        /// <summary>
        /// NetOffice method: dispose instance and all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose event exported proxies with one or more event recipients</param>
        /// <exception cref="COMDisposeException">An unexpected error occurs.</exception>
        public virtual void Dispose(bool disposeEventBinding)
        {
            try
            {
                lock (_disposeLock)
                {
                    // skip check 
                    bool cancel = RaiseOnDispose();
                    if (cancel)
                        return;

                    // in case object export events and 
                    // disposeEventBinding == false we dont remove the object from parents child list
                    bool removeFromParent = true;

                    // set disposing flag
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

                    IEnumerable<ICOMObject> ownerPath = null;
                    if (Factory.HasProxyRemovedRecipients)
                    {
                        ownerPath = Core.GetOwnerPath(this);
                    }

                    // remove himself from parent childlist
                    if ((!Object.ReferenceEquals(_parentObject, null)) && (true == removeFromParent))
                    {
                        _parentObject.RemoveChildObject(this);
                        _parentObject = null;
                    }
                    
                    if (true == removeFromParent)
                    {        
                        // call quit automatically if wanted
                        if (_callQuitInDispose && Settings.EnableAutomaticQuit)
                            new Callers.QuitCaller().TryCall(Settings, Invoker, this);

                        // release proxy
                        ReleaseCOMProxy(ownerPath);

                        // clear supportList reference
                        _listSupportedEntities = null;

                        _isDisposed = true;
                        _isCurrentlyDisposing = false;
                    }
                    else
                        _isCurrentlyDisposing = false;

                }
            }
            catch (Exception exception)
            {
                throw new COMDisposeException("An unexpected error occured while disposing <" + 
                    InstanceName + ">.", exception);
            }
        }
        
        /// <summary>
        /// Call the OnDispose event as service for client callers.
        /// The method implementation ignore any exception in the event handler.
        /// </summary>
        /// <returns>true if cancel is requested</returns>
        protected virtual bool RaiseOnDispose()
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

        #endregion

        #region ICOMObjectTable

        /// <summary>
        /// NetOffice property: returns parent proxy object
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        public ICOMObject ParentObject
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
        /// NetOffice property: Child instances
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false), Category("NetOffice")]
        public IEnumerable<ICOMObject> ChildObjects
        {
            get
            {
                return _listChildObjects;
            }
        }
        /// <summary>
        ///  NetOffice method: add object to child list
        /// </summary>
        /// <param name="childObject">target child instance</param>
        /// <exception cref="COMChildRelationException">Unexpected error</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public void AddChildObject(ICOMObject childObject)
        {
            try
            {
                if (null == childObject)
                    throw new ArgumentNullException("childObject");
                lock (_childListLock)
                {
                    _listChildObjects.Add(childObject);
                }
            }
            catch (Exception exception)
            {
                Console.WriteException(exception);
                throw new COMChildRelationException("Unexpected error while add child instance.", exception);
            }
        }

        /// <summary>
        /// NetOffice property: remove object from child list
        /// </summary>
        /// <param name="childObject">target child instance</param>
        /// <returns>true if childObject has been removed, otherwise false</returns>
        /// <exception cref="COMChildRelationException">Unexpected error</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool RemoveChildObject(ICOMObject childObject)
        {      
            try
            {
                if (null == childObject)
                    throw new ArgumentNullException("childObject");
                lock (_childListLock)
                {
                    return _listChildObjects.Remove(childObject);
                }             
            }
            catch (Exception exception)
            {
                Console.WriteException(exception);
                throw new COMChildRelationException("Unexpected error while remove child instance.", exception);
            }
        }

        #endregion

        #region ICOMObjectTableDisposable
        
        /// <summary>
        /// NetOffice method: dispose all child instances
        /// </summary>
        /// <exception cref="COMDisposeException">An unexpected error occurs.</exception>
        public virtual void DisposeChildInstances()
        {
            DisposeChildInstances(true);
        }

        /// <summary>
        /// NetOffice method: dispose all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose proxies with events and one or more event recipients</param>
        /// <exception cref="COMDisposeException">An unexpected error occurs.</exception>
        public virtual void DisposeChildInstances(bool disposeEventBinding)
        {
            try
            {
                lock (_disposeChildLock)
                {
                    foreach (ICOMObject itemObject in _listChildObjects.ToArray())
                    {                       
                        itemObject.Dispose(disposeEventBinding);
                    }
                    _listChildObjects.Clear();
                }
            }
            catch (Exception exception)
            {
                throw new COMDisposeException("Unexpected error while dispose child instances.", exception);
            }
        }

        #endregion

        #region ICOMObjectEvents
       
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

        #region ICOMObjectAvaility

        /// <summary>
        /// NetOffice method: Returns information the proxy provides a method or property.
        /// Check want be made at runtime through IDispatch interface.
        /// </summary>
        /// <param name="name">name of the enitity</param>
        /// <returns>true if available, otherwise false</returns>
        /// <exception cref="AvailityException">Unexpected error, see inner exception(s) for details.</exception>
        public bool EntityIsAvailable(string name)
        {
            return EntityIsAvailable(name, SupportedEntityType.Both);
        }

        /// <summary>
        /// NetOffice method: Returns information the proxy provides a method or property.
        /// Check want be made at runtime through IDispatch interface.
        /// </summary>
        /// <param name="name">name of the enitity</param>
        /// <param name="searchType">indicate the kind of enitity the caller is looking for</param>
        /// <returns>true if available, otherwise false</returns>
        /// <exception cref="AvailityException">Unexpected error, see inner exception(s) for details.</exception>
        public bool EntityIsAvailable(string name, SupportedEntityType searchType)
        {
            return new SupportedEntityFinder().Find(Factory, ref _listSupportedEntities, searchType, UnderlyingObject, name);
        }

        #endregion

        #region ICOMProxyShareProvider

        /// <summary>
        /// NetOffice method: Returns the inner proxy shared access handler
        /// </summary>
        /// <returns>shared proxy</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        COMProxyShare ICOMProxyShareProvider.GetProxyShare()
        {
            return _proxyShare;
        }

        /// <summary>
        /// NetOffice method: Set the inner proxy shared access handler.
        /// The method want aquire the share 1x times
        /// </summary>
        /// <param name="share">target share</param>
        /// <exception cref="ArgumentNullException">Throws when given share is null(Nothing in Visual Basic)</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        void ICOMProxyShareProvider.SetProxyShare(COMProxyShare share)
        {
            if (null == share)
                throw new ArgumentNullException("share");
            _proxyShare = share;
            _proxyShare.Acquire();
        }

        #endregion

        #region ICloneable

        /// <summary>
        /// NetOffice method: Creates a new object that is a copy of the current instance.
        /// </summary>
        /// <returns>a new object that is a copy of this instance</returns>
        /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public virtual object Clone()
        {
            try
            {
                ICOMObject clone = Activator.CreateInstance(InstanceType, new object[] { Factory, ParentObject, UnderlyingObject }) as ICOMObject;
                ICOMProxyShareProvider shareProvider = clone as ICOMProxyShareProvider;
                shareProvider.SetProxyShare(_proxyShare);
                return clone;
            }
            catch (Exception exception)
            {
                throw new CloneException(exception);
            }
        }

        #endregion

        #region Overrides

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
        /// <param name="obj">target instance to compare</param>
        /// <returns>true if equal, otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool EqualsOnServer(ICOMObject obj)
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

        ///// <summary>
        ///// Determines whether two COMObject instances are equal.
        ///// </summary>
        ///// <param name="objectA"></param>
        ///// <param name="objectB"></param>
        ///// <returns></returns>
        //public static bool operator ==(COMObject objectA, COMDynamicObject objectB)
        //{
        //    if (!Settings.Default.EnableOperatorOverlads)
        //        return Object.ReferenceEquals(objectA, objectB);

        //    if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
        //        return true;
        //    else if (!Object.ReferenceEquals(objectA, null))
        //        return objectA.EqualsOnServer(objectB);
        //    else
        //        return false;
        //}

        ///// <summary>
        ///// Determines whether two COMObject instances are not equal.
        ///// </summary>
        ///// <param name="objectA">first instance</param>
        ///// <param name="objectB">second instance</param>
        ///// <returns>true if equal, otherwise false</returns>
        //public static bool operator !=(COMObject objectA, COMDynamicObject objectB)
        //{
        //    if (!Settings.Default.EnableOperatorOverlads)
        //        return Object.ReferenceEquals(objectA, objectB);

        //    if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
        //        return false;
        //    else if (!Object.ReferenceEquals(objectA, null))
        //        return !objectA.EqualsOnServer(objectB);
        //    else
        //        return true;
        //}


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
            {
                COMObject a = (objectA as COMObject);
                if (null != a)
                    return a.EqualsOnServer(objectB);
                else
                    return false;
            }
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
            {
                COMObject a = objectA as COMObject;
                if (null != a)
                    return !(objectA as COMObject).EqualsOnServer(objectB);
                else
                    return null == objectB ? false : true;
            }
            else
                return true;
        }


        #endregion
    }
}