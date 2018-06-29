using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Reflection;
using System.Collections.Generic;
using System.ComponentModel;
using COMTypes = System.Runtime.InteropServices.ComTypes;
using System.Dynamic;
using System.Collections;
using System.Linq;
using System.Linq.Expressions;
using NetOffice.Resolver;
using NetOffice.Availity;
using NetOffice.Attributes;
using NetOffice.Dynamics;
using NetOffice.Exceptions;
using NetOffice.CoreServices;

namespace NetOffice
{
    /*
        This is designed to use as dynamic in C# or as object in visual basic.
        Allows to use dynamic late-binding with proxy managed service from Netoffice.(best of both worlds)

        NetOffice.Settings.EnableDynamicObjects(currently true by default - since Netoffice 1.7.4.1) want enable
        the behavior that Netoffice returns a COMDynamicObject instance if its
        failed to resolve a wrapper class for a com proxy.

        See tutorials for further informations.
    */

    /// <summary>
    /// Represents a managed COM proxy with dynamic runtime type informations.
    /// </summary>
    [DebuggerDisplay("{InstanceFriendlyName}")]
    [TypeConverter(typeof(Converter.COMDynamicObjectExpandableObjectConverter))]
    public class COMDynamicObject : DynamicObject, ICOMObject, ICOMProxyShareProvider
    {
        #region Nested

        /// <summary>
        /// Plain IEnumerable wrapper implementation
        /// </summary>
        internal class Enumerator : System.Collections.IEnumerable
        {
            private System.Collections.IEnumerator _enumerator;

            internal Enumerator(System.Collections.IEnumerator enumerator)
            {
                _enumerator = enumerator;
            }

            public IEnumerator GetEnumerator()
            {
                return _enumerator;
            }
        }

        /// <summary>
        /// Indicates the COM proxy offers a default property (like this[int index])
        /// </summary>
        private enum DefaultItemSupport
        {
            /// <summary>
            /// No default property available
            /// </summary>
            NoDefaultItem = 0,

            /// <summary>
            /// Default property available as _Default property
            /// </summary>
            PropertyDefault = 1,

            /// <summary>
            /// Default property available as _Default method
            /// </summary>
            MethodDefault = 2,

            /// <summary>
            /// Default property available as Item property
            /// </summary>
            PropertyItem = 3,

            /// <summary>
            /// Default property available as Item method
            /// </summary>
            MethodItem = 4
        }

        /// <summary>
        /// Indicates the COM proxy offers an enumerator
        /// </summary>
        private enum EnumeratorSupport
        {
            /// <summary>
            /// No enumerator available
            /// </summary>
            NoEnumerator = 0,

            /// <summary>
            /// Enumerator available as a property
            /// </summary>
            PropertyEnumerator = 1,

            /// <summary>
            /// Enumerator available as a method
            /// </summary>
            MethodEnumerator = 2
        }

        #endregion

        #region Fields

        /// <summary>
        /// The well know IUnknown Interface ID
        /// </summary>
        private static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        /// <summary>
        /// returns parent instance
        /// </summary>
        protected internal ICOMObject _parentObject;

        /// <summary>
        /// Child instance List
        /// </summary>
        protected internal List<ICOMObject> _listChildObjects = new List<ICOMObject>();

        /// <summary>
        /// Returns instance is currently in disposing progress
        /// </summary>
        protected internal volatile bool _isCurrentlyDisposing;

        /// <summary>
        /// Returns instance is diposed means unusable
        /// </summary>
        protected internal volatile bool _isDisposed;

        /// <summary>
        /// try to call quit in dispose. must be set in top class
        /// </summary>
        protected internal bool _callQuitInDispose;

        /// <summary>
        /// Runtime self description
        /// </summary>
        protected internal DynamicObjectEntity[] _entities;

        /// <summary>
        /// List of runtime supported entities
        /// </summary>
        private Dictionary<string, string> _listSupportedEntities;

        /// <summary>
        /// Returns a shared access wrapper arrount the native wrapped proxy
        /// </summary>
        protected internal COMProxyShare _proxyShare;

        /// <summary>
        /// Monitor lock object for accessing the child list
        /// </summary>
        private object _childListLock = new object();

        /// <summary>
        /// monitor lock object for accessing the child list
        /// </summary>
        private object _disposeChildLock = new object();

        /// <summary>
        /// Monitor lock object for the main dispose method
        /// </summary>
        private object _disposeLock = new object();

        /// <summary>
        /// Indicates the instance offers an enumerator
        /// </summary>
        private EnumeratorSupport _enumerator;

        /// <summary>
        /// Indicates the instance offers an default property
        /// </summary>
        private DefaultItemSupport _defaultItem;

        /// <summary>
        /// CheckEntities Monitor Lock
        /// </summary>
        private object _entitiesLock = new object();

        /// <summary>
        /// Empty arguments dumy
        /// </summary>
        private static object[] _emptyArgs = new object[0];

        /// <summary>
        /// Self Type Cache
        /// </summary>
        private static Type _instanceType = typeof(COMDynamicObject);

        /// <summary>
        /// Contract Type cache field
        /// </summary>
        private Type _contractType = null;

        /// <summary>
        /// Given ProgID in ctor or null
        /// </summary>
        private string _progId;

        /// <summary>
        /// Dynamic accessible instance members
        /// </summary>
        private static string[] _selfDynamicMemberNames;

        /// <summary>
        /// Invalid proxy error message
        /// </summary>
        private static string _invalidComProxy = "Given argument isn't a com proxy.";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates new instance with given proxy
        /// </summary>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        public COMDynamicObject(object comProxy)
        {
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");
            if (!(comProxy is MarshalByRefObject))
                throw new ArgumentException(_invalidComProxy);
            Factory = Core.Default;
            SyncRoot = new object();
            ParentObject = null;
            _proxyShare = Factory.InternalObjectActivator.CreateNewProxyShare(this, comProxy);
            UnderlyingType = comProxy.GetType();
            Factory.InternalObjectRegister.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();
            Factory.CheckInitialize();
        }

        /// <summary>
        /// Creates new instance with given proxy and parent info
        /// </summary>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        public COMDynamicObject(ICOMObject parentObject, object comProxy)
        {
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");
            if (!(comProxy is MarshalByRefObject))
                throw new ArgumentException(_invalidComProxy);

            if (null != parentObject)
                Factory = parentObject.Factory;
            else
                Factory = Core.Default;
            SyncRoot = new object();

            ParentObject = parentObject;
            _proxyShare = Factory.InternalObjectActivator.CreateNewProxyShare(this, comProxy);
            UnderlyingType = comProxy.GetType();
            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                ParentObject.AddChildObject(this);

            Factory.InternalObjectRegister.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();
            Factory.CheckInitialize();
        }

        /// <summary>
        /// Creates new (root) instance with given managed proxy
        /// </summary>
        /// <param name="comObject">managed proxy</param>
        public COMDynamicObject(ICOMObject comObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            Factory = comObject.Factory;
            SyncRoot = new object();

            ICOMProxyShareProvider shareProvider = comObject as ICOMProxyShareProvider;
            if (null != shareProvider)
                _proxyShare = shareProvider.GetProxyShare();
            else
                _proxyShare = Factory.InternalObjectActivator.CreateNewProxyShare(this, comObject.UnderlyingObject);

            UnderlyingType = comObject.UnderlyingType;

            Factory.InternalObjectRegister.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();
            Factory.CheckInitialize();
        }

        /// <summary>
        /// Creates new instance with given proxy and parent info
        /// </summary>
        /// <param name="factory">current factory instance or null for defauslt</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        public COMDynamicObject(Core factory, ICOMObject parentObject, object comProxy)
        {
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");
            if (!(comProxy is MarshalByRefObject))
                throw new ArgumentException(_invalidComProxy);

            if (null == factory)
                factory = Core.Default;
            Factory = factory;
            SyncRoot = new object();

            ParentObject = parentObject;
            _proxyShare = Factory.InternalObjectActivator.CreateNewProxyShare(this, comProxy);

            UnderlyingType = comProxy.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                ParentObject.AddChildObject(this);

            Factory.InternalObjectRegister.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();
            Factory.CheckInitialize();
        }

        /// <summary>
        /// Creates new instance with given proxy and parent info
        /// </summary>
        /// <param name="factory">current factory instance or null for defauslt</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">proxy share instead of proxy</param>
        public COMDynamicObject(Core factory, ICOMObject parentObject, COMProxyShare comProxy)
        {
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");

            if (null == factory)
                factory = Core.Default;
            Factory = factory;
            SyncRoot = new object();

            ParentObject = parentObject;
            _proxyShare = comProxy;

            UnderlyingType = _proxyShare.Proxy.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                ParentObject.AddChildObject(this);

            Factory.InternalObjectRegister.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();
            Factory.CheckInitialize();
        }

        /// <summary>
        /// Create new instance from given progid
        /// </summary>
        /// <param name="factory">used factory core</param>
        /// <param name="progId">progid as any</param>
        public COMDynamicObject(Core factory, string progId)
        {
            if (String.IsNullOrEmpty(progId))
                throw new ArgumentNullException("progId");

            object underlyingObject = CreateFromProgId(progId);
            _proxyShare = Factory.InternalObjectActivator.CreateNewProxyShare(this, underlyingObject);

            SyncRoot = new object();

            Factory = null != factory ? factory : Core.Default;
            Factory.InternalObjectRegister.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();

            _progId = progId;

            Factory.CheckInitialize();
        }

        /// <summary>
        /// Create new instance from given progid
        /// </summary>
        /// <param name="progId">prog id as any</param>
        public COMDynamicObject(string progId)
        {
            if (String.IsNullOrEmpty(progId))
                throw new ArgumentNullException("progId");

            Factory = Core.Default;
            SyncRoot = new object();

            object underlyingObject = CreateFromProgId(progId);
            _proxyShare = Factory.InternalObjectActivator.CreateNewProxyShare(this, underlyingObject);

            Factory.InternalObjectRegister.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();

            _progId = progId;

            Factory.CheckInitialize();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Return Value in TryConvert if no conversion is available.
        /// False may cause an exception from the current language service,
        /// otherwise the conversion result is just null(Nothing in Visual Basic)
        /// Default: false
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public static bool TryConvertFailResult { get; set; }

        /// <summary>
        /// Dynamic accessible instance members
        /// </summary>
        private static string[] SelfDynamicMemberNames
        {
            get
            {
                if (null == _selfDynamicMemberNames)
                {
                    List<string> list = new List<string>();
                    list.Add("Dispose");
                    list.Add("DisposeChildInstances");
                    _selfDynamicMemberNames = list.ToArray();
                }
                return _selfDynamicMemberNames;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Create a COMDynamicObject shallow copy from COMObject instance.
        /// The shallow copy is a root instance in com proxy management without child instances.
        /// Given COMObject instance and shallow copy share the same proxy.
        /// </summary>
        /// <param name="comObject">COMObject instance</param>
        /// <returns>COMDynamicObject shallow copy</returns>
        /// <exception cref="ArgumentNullException">throws when comObject is null</exception>
        public static COMDynamicObject ConvertTo(ICOMObject comObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            return new COMDynamicObject(comObject.UnderlyingObject);
        }

        /// <summary>
        /// Release com proxy
        /// </summary>
        private void ReleaseCOMProxy(IEnumerable<ICOMObject> ownerPath, bool isRootObject = false)
        {
            // release himself from COM Runtime System
            if (!_proxyShare.Released)
            {
                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(ContractType.Namespace, ContractType.Name, "NetOffice::ReleaseCOMProxy", PerformanceTrace.CallType.Method);
                _proxyShare.Release();
                Factory.InternalObjectRegister.RemoveObjectFromList(this, ownerPath);
                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(ContractType.Namespace, ContractType.Name, "NetOffice::ReleaseCOMProxy");
            }
        }

        /// <summary>
        /// Calls the OnDispose event as service for client callers
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

        /// <summary>
        /// Check for GetEntites has been called for the instance and call if not
        /// </summary>
        private void CheckEntities()
        {
            lock (_entitiesLock)
            {
                if (null == _entities)
                    _entities = GetEntities();
            }
        }

        /// <summary>
        /// Creates underlying type and underlying object from given prog id
        /// </summary>
        /// <param name="progId">progid as any</param>
        /// <returns>newly created instance</returns>
        private object CreateFromProgId(string progId)
        {
            bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(ContractType.Namespace, ContractType.Name, "NetOffice::CreateFromProgId", PerformanceTrace.CallType.Method);

            UnderlyingType = System.Type.GetTypeFromProgID(progId, true);
            object underlyingObject = Activator.CreateInstance(UnderlyingType);

            if (measureStarted)
                Settings.PerformanceTrace.StopMeasureTime(ContractType.Namespace, ContractType.Name, "NetOffice::CreateFromProgId");

            return underlyingObject;
        }

        /// <summary>
        /// Recieve self description from UnderlyingObject through IDispatch
        /// </summary>
        /// <returns>entity collection</returns>
        private DynamicObjectEntity[] GetEntities()
        {
            List<DynamicObjectEntity> result = new List<DynamicObjectEntity>();

            IDispatch dispatch = UnderlyingObject as IDispatch;
            if (null == dispatch)
                throw new COMException("Unable to cast IDispatch.");

            COMTypes.ITypeInfo typeInfo = dispatch.GetTypeInfo(0, 0);
            if (null == typeInfo)
                throw new COMException("Unable to get type informations.");

            IntPtr typeAttrPointer;
            typeInfo.GetTypeAttr(out typeAttrPointer);

            COMTypes.TYPEATTR typeAttr = (COMTypes.TYPEATTR)Marshal.PtrToStructure(typeAttrPointer, typeof(COMTypes.TYPEATTR));
            for (int i = 0; i < typeAttr.cFuncs; i++)
            {
                string entityName, entityDescription, entityHelpFilePath;
                int entityHelpContext;
                IntPtr funcDescPointer = IntPtr.Zero;
                COMTypes.FUNCDESC funcDesc;
                typeInfo.GetFuncDesc(i, out funcDescPointer);
                funcDesc = (COMTypes.FUNCDESC)Marshal.PtrToStructure(funcDescPointer, typeof(COMTypes.FUNCDESC));

                if (funcDesc.funckind == COMTypes.FUNCKIND.FUNC_DISPATCH)
                {
                    typeInfo.GetDocumentation(funcDesc.memid, out entityName, out entityDescription, out entityHelpContext, out entityHelpFilePath);
                    CheckEnumeratorEntity(entityName, funcDesc.invkind);
                    CheckDefaultEntity(entityName, funcDesc.invkind);

                    switch (funcDesc.invkind)
                    {
                        case COMTypes.INVOKEKIND.INVOKE_PROPERTYGET:
                            {
                                DynamicObjectEntity writeProperty = FindInCollection(result, entityName, DynamicObjectEntity.EntityKind.PropertyWritable);
                                if (null != writeProperty)
                                    result.Add(new DynamicObjectEntity(entityName, DynamicObjectEntity.EntityKind.PropertyReadonly));
                                break;
                            }
                        case COMTypes.INVOKEKIND.INVOKE_PROPERTYPUT:
                        case COMTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF:
                            {
                                DynamicObjectEntity readProperty = FindInCollection(result, entityName, DynamicObjectEntity.EntityKind.PropertyReadonly);
                                if (null != readProperty)
                                    readProperty.Kind = DynamicObjectEntity.EntityKind.PropertyWritable;
                                else
                                    result.Add(new DynamicObjectEntity(entityName, DynamicObjectEntity.EntityKind.PropertyWritable));
                                break;
                            }
                        case COMTypes.INVOKEKIND.INVOKE_FUNC:
                            {
                                result.Add(new DynamicObjectEntity(entityName, DynamicObjectEntity.EntityKind.Method));
                                break;
                            }
                    }
                }
                typeInfo.ReleaseFuncDesc(funcDescPointer);
            }

            return result.ToArray();
        }

        /// <summary>
        /// Check and stores the information the given proxy entity is an enumerator
        /// </summary>
        /// <param name="name">name of the entity</param>
        /// <param name="kind">kind of the entity</param>
        private void CheckEnumeratorEntity(string name, COMTypes.INVOKEKIND kind)
        {
            if (name != "_NewEnum")
                return;
            switch (kind)
            {
                case COMTypes.INVOKEKIND.INVOKE_FUNC:
                    _enumerator = EnumeratorSupport.MethodEnumerator;
                    break;
                default:
                    _enumerator = EnumeratorSupport.PropertyEnumerator;
                    break;
            }
        }

        /// <summary>
        /// Check and stores the information the given proxy entity is a default property
        /// </summary>
        /// <param name="name">name of the entity</param>
        /// <param name="kind">kind of the entity</param>
        private void CheckDefaultEntity(string name, COMTypes.INVOKEKIND kind)
        {
            if (name == "_Default")
            {
                switch (kind)
                {
                    case COMTypes.INVOKEKIND.INVOKE_FUNC:
                        _defaultItem = DefaultItemSupport.MethodDefault;
                        break;
                    default:
                        _defaultItem = DefaultItemSupport.PropertyDefault;
                        break;
                }
            }
            else if (name == "Item")
            {
                switch (kind)
                {
                    case COMTypes.INVOKEKIND.INVOKE_FUNC:
                        _defaultItem = DefaultItemSupport.MethodItem;
                        break;
                    default:
                        _defaultItem = DefaultItemSupport.PropertyItem;
                        break;
                }
            }
        }

        /// <summary>
        /// Find item in collection. (Wrapper to bypass missing Linq in former .Net runtimes)
        /// </summary>
        /// <param name="values">collection</param>
        /// <param name="name">target name</param>
        /// <param name="kind">target kind</param>
        /// <returns>item or null</returns>
        private DynamicObjectEntity FindInCollection(IEnumerable<DynamicObjectEntity> values, string name, DynamicObjectEntity.EntityKind kind)
        {
            foreach (DynamicObjectEntity item in values)
            {
                if (item.Name == name && item.Kind == kind)
                    return item;
            }
            return null;
        }

        /// <summary>
        /// Check binder want an enumerator
        /// </summary>
        /// <param name="binder">given binder</param>
        /// <returns>true if binder want an enumerator, otherwise false</returns>
        private bool IsEnumeratorBinder(GetMemberBinder binder)
        {
            return null != binder ? binder.Name == "_NewEnum" : false;
        }

        /// <summary>
        /// Check binder want an enumerator
        /// </summary>
        /// <param name="binder">given binder</param>
        /// <returns>true if binder want an enumerator, otherwise false</returns>
        private bool IsEnumeratorBinder(InvokeMemberBinder binder)
        {
            return null != binder ? binder.Name == "_NewEnum" : false;
        }

        /// <summary>
        /// Invoke the proxy enumerator
        /// </summary>
        /// <returns>IEnumerable sequence</returns>
        private System.Collections.IEnumerable InvokeEnumerator()
        {
            CheckEntities();

            switch (_enumerator)
            {
                case EnumeratorSupport.PropertyEnumerator:
                    return new Enumerator(NetOffice.Utils.GetVariantEnumeratorAsProperty(this, true));
                case EnumeratorSupport.MethodEnumerator:
                    return new Enumerator(NetOffice.Utils.GetVariantEnumeratorAsMethod(this, true));
                default:
                    return null;
            }
        }

        /// <summary>
        /// Invoke a proxy method
        /// </summary>
        /// <param name="name">method name</param>
        /// <returns>return value or null</returns>
        private object InvokeMethod(string name)
        {
            if (IsSelfDynamicMemberName(name))
                return InstanceType.InvokeMember(name, System.Reflection.BindingFlags.InvokeMethod, null, this, new object[0]);

            object returnItem = Invoker.MethodReturn(this, name);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem, true);
                return newObject;
            }
            else
            {
                return returnItem;
            }
        }

        /// <summary>
        /// Invoke a proxy method
        /// </summary>
        /// <param name="name">method name</param>
        /// <param name="args">method arguments</param>
        /// <returns>return value or null</returns>
        private object InvokeMethod(string name, object[] args)
        {
            if (IsSelfDynamicMemberName(name))
                return InstanceType.InvokeMember(name, System.Reflection.BindingFlags.InvokeMethod, null, this, args);

            args = Invoker.ValidateParamsArray(args);
            object returnItem = Invoker.MethodReturn(this, name, args);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem, true);
                return newObject;
            }
            else
            {
                return returnItem;
            }
        }

        /// <summary>
        /// Invoke a proxy method
        /// </summary>
        /// <param name="name">method name</param>
        /// <param name="args">method arguments</param>
        /// <param name="value">additional argument</param>
        /// <returns>return value or null</returns>
        private object InvokeMethod(string name, object[] args, object value)
        {
            int i = 0;
            object[] arguments = new object[args.Length + 1];
            foreach (var item in args)
            {
                arguments[i] = item;
                i++;
            }
            arguments[i] = value;

            args = Invoker.ValidateParamsArray(arguments);
            object returnItem = Invoker.MethodReturn(this, name, args);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem, true);
                return newObject;
            }
            else
            {
                return returnItem;
            }
        }

        /// <summary>
        /// Invoke a proxy property for read access
        /// </summary>
        /// <param name="name">property name</param>
        /// <returns>property value</returns>
        private object InvokePropertyGet(string name)
        {
            object returnItem = Invoker.PropertyGet(this, name);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem, true);
                return newObject;
            }
            else
            {
                return returnItem;
            }
        }

        /// <summary>
        /// Invoke a proxy property for read access
        /// </summary>
        /// <param name="name">property name</param>
        /// <param name="args">arguments</param>
        /// <returns>property value</returns>
        private object InvokePropertyGet(string name, object[] args)
        {
            args = Invoker.ValidateParamsArray(args);
            object returnItem = Invoker.PropertyGet(this, name, args);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem, true);
                return newObject;
            }
            else
            {
                return returnItem;
            }
        }

        /// <summary>
        /// Invoke a proxy property for write access
        /// </summary>
        /// <param name="name">property name</param>
        /// <param name="args">arguments</param>
        private void InvokePropertySet(string name, object[] args)
        {
            args = Invoker.ValidateParamsArray(args);
            Invoker.PropertySet(this, name, args );
        }

        /// <summary>
        /// Invoke a proxy property for write access
        /// </summary>
        /// <param name="name">property name</param>
        /// <param name="args">arguments</param>
        /// <param name="value">additional argument</param>
        private void InvokePropertySet(string name, object[] args, object value)
        {
            int i = 0;
            object[] arguments = new object[args.Length + 1];
            foreach (var item in args)
            {
                arguments[i] = item;
                i++;
            }
            arguments[i] = value;
            args = Invoker.ValidateParamsArray(arguments);
            Invoker.PropertySet(this, name, args);
        }

        /// <summary>
        /// DNUL for compatibility
        /// </summary>
        /// <param name="name">member name</param>
        /// <returns>true if name match, otherwise false</returns>
        private bool IsSelfDynamicMemberName(string name)
        {
            foreach (var item in SelfDynamicMemberNames)
            {
                if (item == name)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// DNUL for compatibility
        /// </summary>
        /// <returns>true if proxy has quit method, otherwise false</returns>
        private bool HasQuitMethod()
        {
            CheckEntities();
            if (null == _entities || _entities.Length == 0)
                return false;
            foreach (var item in _entities)
            {
                if (item.Kind == DynamicObjectEntity.EntityKind.Method && item.Name == "Quit")
                    return true;
            }
            return false;
        }

        #endregion

        #region ICOMObject

        /// <summary>
        /// Monitor Lock
        /// </summary>
        public object SyncRoot { get; private set; }

        /// <summary>
        /// The associated factory
        /// </summary>
        public Core Factory { get; private set; }

        /// <summary>
        /// The associated invoker
        /// </summary>
        public Invoker Invoker
        {
            get
            {
                if (null != Factory)
                    return Factory.Invoker;
                else
                    return Invoker.Default;
            }
        }

        /// <summary>
        /// The associated console
        /// </summary>
        public DebugConsole Console
        {
            get
            {
                if (null != Factory)
                    return Factory.Console;
                else
                    return DebugConsole.Default;
            }
        }

        /// <summary>
        /// The associated settings
        /// </summary>
        public Settings Settings
        {
            get
            {
                if (null != Factory)
                    return Factory.Settings;
                else
                    return Settings.Default;
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
                ICOMObject clone = (ICOMObject)Activator.CreateInstance(typeof(T), new object[] { Factory, ParentObject, UnderlyingObject });

                ICOMProxyShareProvider shareProvider = clone as ICOMProxyShareProvider;
                if (null == shareProvider)
                    throw new InvalidCastException("Newly created instance does not implement the ICOMProxyShareProvider interface.");
                shareProvider.SetProxyShare(_proxyShare);

                IAutomaticQuit quitObject = clone as IAutomaticQuit;
                if (null != quitObject)
                    quitObject.Enabled = false;

                return clone as T;
            }
            catch (Exception exception)
            {
                throw new CloneException(exception);
            }
        }

        /// <summary>
        /// Determines whether two ICOMObject instances pointing to the same remote server instance.
        /// </summary>
        /// <param name="obj">target instance to compare</param>
        /// <returns>true if equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool EqualsOnServer(object obj)
        {
            return EqualsOnServer(obj as ICOMObject);
        }

        #endregion

        #region ICOMObjectProxy

        /// <summary>
        /// Returns the native wrapped proxy
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        public object UnderlyingObject
        {
            get
            {
                return _proxyShare.Proxy;
            }
        }

        /// <summary>
        /// Returns Type of native proxy
        /// </summary>
        public Type UnderlyingType { get; private set; }

        /// <summary>
        /// Friendly Name of the NetOffice Wrapper class
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        public string InstanceFriendlyName
        {
            get
            {
                if (null != _progId)
                    return "Dynamic(" + _progId + ")";
                else
                    return "Dynamic(" + InstanceFriendlyName + ")";
            }
        }

        /// <summary>
        /// Name of the hosting NetOffice component
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        public string InstanceComponentName
        {
            get
            {
                return "NetOffice.Core";
            }
        }

        /// <summary>
        /// Type informations from ICOMObject instance
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        public Type InstanceType
        {
            get
            {
                return _instanceType;
            }
        }

        /// <summary>
        /// Type informations from ICOMObject contract
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual Type ContractType
        {
            get
            {
                if (null == _contractType)
                {
                    Type[] allInterfaces = InstanceType.GetInterfaces();
                    _contractType = allInterfaces.Except(allInterfaces.SelectMany(t => t.GetInterfaces())).FirstOrDefault(e => e.HasCustomAttribute<TypeIdAttribute>());
                }
                return _contractType;
            }
        }

        #endregion

        #region ICOMObjectDisposable

        /// <summary>
        /// These event was called from Dispose and you can skip the dipose operation here if you want. the event can be helpful for troubleshooting if you dont know why your objects beeing disposed
        /// </summary>
        public event OnDisposeEventHandler OnDispose;

        /// <summary>
        /// Returns instance is already diposed
        /// </summary>
        public bool IsDisposed
        {
            get
            {
                return _isDisposed;
            }
        }

        /// <summary>
        /// Returns instance is currently in diposing progress
        /// </summary>
        public bool IsCurrentlyDisposing
        {
            get
            {
                return _isCurrentlyDisposing;
            }
        }

        /// <summary>
        /// Dispose instance and all child instances
        /// </summary>
        /// <exception cref="COMDisposeException">An unexpected error occurs.</exception>
        public virtual void Dispose()
        {
            Dispose(true);
        }

        /// <summary>
        /// Dispose instance and all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose proxies with events and one or more event recipients</param>
        /// <exception cref="COMDisposeException">An unexpected error occurs.</exception>
        public virtual void Dispose(bool disposeEventBinding)
        {
            bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(ContractType.Namespace, ContractType.Name, "NetOffice::Dispose", PerformanceTrace.CallType.Method);
            bool isRootObject = null == ParentObject;
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
                    if (Factory.InternalObjectRegister.HasRemovedRecipients)
                    {
                        ownerPath = NetOffice.CoreServices.Internal.CoreManagement.GetOwnerPath(this);
                    }

                    // remove himself from parent childlist
                    if ((!Object.ReferenceEquals(ParentObject, null)) && (true == removeFromParent))
                    {
                        ParentObject.RemoveChildObject(this);
                        ParentObject = null;
                    }

                    if (true == removeFromParent)
                    {
                        // call quit automatically if wanted
                        if (Settings.EnableAutomaticQuit && HasQuitMethod())
                            new Callers.QuitCaller().TryCall(Settings, Invoker, this);

                        // release proxy
                        ReleaseCOMProxy(ownerPath, isRootObject);

                        // clear supportList reference
                        _listSupportedEntities = null;

                        _isDisposed = true;
                        _isCurrentlyDisposing = false;
                    }
                    else
                        _isCurrentlyDisposing = false;
                }

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(ContractType.Namespace, ContractType.Name, "NetOffice::Dispose");
            }
            catch (Exception exception)
            {
                throw new COMDisposeException("An unexpected error occured while disposing <" +
                    InstanceType.FullName + ">.", exception);
            }
        }

        #endregion

        #region ICOMObjectTable

        /// <summary>
        /// Returns parent proxy object
        /// </summary>
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
        /// Child instances
        /// </summary>
        public IEnumerable<ICOMObject> ChildObjects
        {
            get
            {
                return _listChildObjects;
            }
        }

        /// <summary>
        /// Add object to child list
        /// </summary>
        /// <param name="childObject">>target child instance</param>
        /// <exception cref="COMChildRelationException">Unexpected error</exception>
        public void AddChildObject(ICOMObject childObject)
        {
            try
            {
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
        /// Remove object from child list
        /// </summary>
        /// <param name="childObject">target child instance</param>
        /// <exception cref="COMChildRelationException">Unexpected error</exception>
        public bool RemoveChildObject(ICOMObject childObject)
        {
            try
            {
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
        /// Dispose all child instances
        /// </summary>
        /// <exception cref="COMDisposeException">An unexpected error occurs.</exception>
        public virtual void DisposeChildInstances()
        {
            DisposeChildInstances(true);
        }

        /// <summary>
        /// Dispose all child instances
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

        /// <summary>
        /// Removes an instance from its current position in com proxy management and make him a root object
        /// </summary>
        /// <typeparam name="T">cast instance into result type</typeparam>
        /// <returns>instance result as a root proxy</returns>
        /// <exception cref="CreateInstanceException">Unexpected error</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        public T TakeObject<T>() where T : class, ICOMObject
        {
            try
            {
                var parentObject = ParentObject;
                if (null != parentObject)
                {
                    parentObject.RemoveChildObject(this);
                }

                return Activator.CreateInstance(typeof(T), Factory, null, UnderlyingObject) as T;
            }
            catch (Exception exception)
            {
                throw new CreateInstanceException(exception);
            }
        }

        #endregion

        #region ICOMObjectEvents

        /// <summary>
        /// Unsupported
        /// </summary>
        public bool IsEventBinding
        {
            get
            {   // unsupported in dynamics
                return false;
            }
        }

        /// <summary>
        /// Unsupported
        /// </summary>
        public bool IsEventBridgeInitialized
        {
            get
            {   // unsupported in dynamics
                return false;
            }
        }

        /// <summary>
        /// Unsupported
        /// </summary>
        public bool IsWithEventRecipients
        {
            get
            {
                // unsupported in dynamics
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
        /// Creates a new object that is a copy of the current instance.
        /// </summary>
        /// <returns>a new object that is a copy of this instance</returns>
        /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public virtual object Clone()
        {
            try
            {
                ICOMObject clone = (ICOMObject)Activator.CreateInstance(InstanceType, new object[] { Factory, ParentObject, UnderlyingObject });

                ICOMProxyShareProvider shareProvider = clone as ICOMProxyShareProvider;
                if (null == shareProvider)
                    throw new InvalidCastException("Newly created instance does not implement the ICOMProxyShareProvider interface.");
                shareProvider.SetProxyShare(_proxyShare);

                IAutomaticQuit quitObject = clone as IAutomaticQuit;
                if (null != quitObject)
                    quitObject.Enabled = false;

                return clone;
            }
            catch (Exception exception)
            {
                throw new CloneException(exception);
            }
        }

        #endregion

        #region Equals

        /// <summary>
        /// Determines whether two ICOMObject instances pointing to the same remote server instance.
        /// </summary>
        /// <param name="objectA">first instance to compare</param>
        /// <param name="objectB">second instance to compare</param>
        /// <returns>true if arguments equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        public static bool EqualsOnServer(object objectA, object objectB)
        {
            ICOMObject objA = objectA as ICOMObject;
            ICOMObject objB = objectA as ICOMObject;

            if (null != objA)
                return objA.EqualsOnServer(objB);
            else if (null != objB)
                return false;
            else
                return Object.ReferenceEquals(objA, objectB);
        }

        /// <summary>
        /// Determines whether two ICOMObject instances pointing to the same remote server instance.
        /// </summary>
        /// <param name="obj">target instance to compare</param>
        /// <returns>true if equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
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
                throw new NetOfficeCOMException("Unexpected error during semantically instance comparsion.", exception);
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

        #endregion

        #region Operators

        /// <summary>
        /// Determines whether two COMObject instances are equal.
        /// </summary>
        /// <param name="objectA"></param>
        /// <param name="objectB"></param>
        /// <returns>true if arguments equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        public static bool operator ==(COMDynamicObject objectA, ICOMObject objectB)
        {
            if (!Settings.EnableOperatorOverloadsInternal)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return true;
            else if (!Object.ReferenceEquals(objectA, null))
                return objectA.EqualsOnServer(objectB);
            else
                return false;
        }

        /// <summary>
        /// Determines whether two COMObject instances are not equal.
        /// </summary>
        /// <param name="objectA">first instance</param>
        /// <param name="objectB">second instance</param>
        /// <returns>true if arguments equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        public static bool operator !=(COMDynamicObject objectA, ICOMObject objectB)
        {
            if (!Settings.EnableOperatorOverloadsInternal)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return false;
            else if (!Object.ReferenceEquals(objectA, null))
                return !objectA.EqualsOnServer(objectB);
            else
                return true;
        }

        /// <summary>
        /// Determines whether two COMObject instances are equal.
        /// </summary>
        /// <param name="objectA"></param>
        /// <param name="objectB"></param>
        /// <returns>true if arguments equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        public static bool operator ==(COMDynamicObject objectA, COMDynamicObject objectB)
        {
            if (!Settings.EnableOperatorOverloadsInternal)
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
        /// <returns>true if arguments equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        public static bool operator ==(COMDynamicObject objectA, object objectB)
        {
            if (!Settings.EnableOperatorOverloadsInternal)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return true;
            else if (!Object.ReferenceEquals(objectA, null))
                return objectA.EqualsOnServer(objectB as ICOMObject);
            else
                return false;
        }

        /// <summary>
        /// Determines whether two COMObject instances are equal.
        /// </summary>
        /// <param name="objectA">first instance</param>
        /// <param name="objectB">second instance</param>
        /// <returns>true if arguments equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        public static bool operator ==(object objectA, COMDynamicObject objectB)
        {
            if (!Settings.EnableOperatorOverloadsInternal)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return true;
            else if (!Object.ReferenceEquals(objectA, null))
            {
                ICOMObject a = (objectA as ICOMObject);
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
        /// <returns>true if arguments equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        public static bool operator !=(COMDynamicObject objectA, COMDynamicObject objectB)
        {
            if (!Settings.EnableOperatorOverloadsInternal)
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
        /// <returns>true if arguments equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        public static bool operator !=(COMDynamicObject objectA, object objectB)
        {
            if (!Settings.EnableOperatorOverloadsInternal)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return false;
            else if (!Object.ReferenceEquals(objectA, null))
                return !objectA.EqualsOnServer(objectB as ICOMObject);
            else
                return true;
        }

        /// <summary>
        /// Determines whether two COMObject instances are not equal.
        /// </summary>
        /// <param name="objectA">first instance</param>
        /// <param name="objectB">second instance</param>
        /// <returns>true if arguments equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        public static bool operator !=(object objectA, COMDynamicObject objectB)
        {
            if (!Settings.EnableOperatorOverloadsInternal)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return false;
            else if (!Object.ReferenceEquals(objectA, null))
            {
                ICOMObject a = objectA as ICOMObject;
                if (null != a)
                    return !a.EqualsOnServer(objectB);
                else
                    return null == objectB ? false : true;
            }
            else
                return true;
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
        /// Determines whether two Object instances are equal.
        /// </summary>
        /// <returns>true if equal, otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public override bool Equals(Object obj)
        {
            return base.Equals(obj);
        }

        /// <summary>
        /// Provides a System.Dynamic.DynamicMetaObject that dispatches to the dynamic virtual
        /// methods. The object can be encapsulated inside another System.Dynamic.DynamicMetaObject
        /// to provide custom behavior for individual actions. This method supports the Dynamic
        /// Language Runtime infrastructure for language implementers and it is not intended
        /// to be used directly from your code.
        /// </summary>
        /// <param name="parameter">The expression that represents System.Dynamic.DynamicMetaObject to dispatch to the dynamic virtual methods.</param>
        /// <returns> An object of the System.Dynamic.DynamicMetaObject type.</returns>
        public override DynamicMetaObject GetMetaObject(Expression parameter)
        {
            DynamicMetaObject metaObject = base.GetMetaObject(parameter);
            return new COMDynamicMetaObject(metaObject);
        }

        /// <summary>
        /// Returns a sequence of all dynamic member names.
        /// </summary>
        /// <returns>a sequence that contains dynamic member names.</returns>
        public override IEnumerable<string> GetDynamicMemberNames()
        {
            CheckEntities();
            if (null == _entities)
                return new string[0];

            string[] selfMembers = SelfDynamicMemberNames;

            int i = 0;
            string[] names = new string[_entities.Length + selfMembers.Length];
            foreach (DynamicObjectEntity item in _entities)
            {
                names[i] = item.Name;
                i++;
            }

            foreach (var item in selfMembers)
            {
                names[i] = item;
                i++;
            }

            return names;
        }

        /// <summary>
        /// Provides implementation for type conversion operations.
        /// </summary>
        /// <param name="binder">Provides information about the conversion operation.</param>
        /// <param name="result">The result of the type conversion operation.</param>
        /// <returns>true if the operation is successful; otherwise, false. </returns>
        public override bool TryConvert(ConvertBinder binder, out object result)
        {
            // Good to know:
            // Confusing stuff about dynamic and implicit/explicit conversions
            // https://stackoverflow.com/questions/3492955/dynamicobject-tryconvert-not-called-when-casting-to-interface-type
            // Not sure what John Skeet means here to handle that better with IDynamicMetaObjectProvider - i tried his idea and fail

            CheckEntities();

            if (binder.Type == typeof(System.Collections.IEnumerable))
            {
                result = InvokeEnumerator();
                return true;
            }
            else if (binder.Type == typeof(string))
            {
                result = InstanceFriendlyName;
                return true;
            }
            else if (binder.Type == typeof(COMObject))
            {
                result = new COMObject(Factory, ParentObject, _proxyShare);
                return true;
            }
            else
            {
                string className = TypeDescriptor.GetClassName(UnderlyingObject);

                Guid typeId = Guid.Empty;
                Guid componentId = Guid.Empty;

                CoreTypeExtensions.GetComponentAndTypeId(Factory, UnderlyingObject, ref componentId, ref typeId);
                ITypeFactory factoryInfo = CoreFactoryExtensions.GetTypeFactory(Factory, this, UnderlyingObject, componentId, typeId, true);

                if (null != factoryInfo && factoryInfo.ContainsType(binder.ReturnType))
                {
                    string fullClassName = factoryInfo.FactoryName + "." + className;
                    if (fullClassName.Equals(binder.ReturnType.FullName))
                    {
                        ICOMObject instance = Activator.CreateInstance(binder.ReturnType, new object[] { Factory, ParentObject, UnderlyingObject }) as ICOMObject;
                        ICOMProxyShareProvider shareProvider = instance as ICOMProxyShareProvider;
                        if (null != shareProvider)
                            shareProvider.SetProxyShare(_proxyShare);
                        result = instance;
                        return true;
                    }
                    else
                    {
                        result = null;
                        return TryConvertFailResult;
                    }
                }
                else
                {
                    result = null;
                    return TryConvertFailResult;
                }
            }
        }

        /// <summary>
        /// Provides the implementation for operations that get a value by index.
        /// </summary>
        /// <param name="binder">Provides information about the operation.</param>
        /// <param name="indexes">The indexes that are used in the operation.</param>
        /// <param name="result">The result of the index operation.</param>
        /// <returns>true if the operation is successful; otherwise, false.</returns>
        public override bool TryGetIndex(GetIndexBinder binder, object[] indexes, out object result)
        {
            CheckEntities();

            result = null;
            switch (_defaultItem)
            {
                case DefaultItemSupport.PropertyDefault:
                    result = InvokePropertyGet("_Default", indexes);
                    return true;
                case DefaultItemSupport.MethodDefault:
                    result = InvokeMethod("_Default", indexes);
                    return true;
                case DefaultItemSupport.PropertyItem:
                    result = InvokePropertyGet("Item", indexes);
                    return true;
                case DefaultItemSupport.MethodItem:
                    result = InvokeMethod("Item", indexes);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Provides the implementation for operations that set a value by index.
        /// </summary>
        /// <param name="binder">Provides information about the operation.</param>
        /// <param name="indexes">The indexes that are used in the operation.</param>
        /// <param name="value">The value to set to the object that has the specified index.</param>
        /// <returns>true if the operation is successful; otherwise, false.</returns>
        public override bool TrySetIndex(SetIndexBinder binder, object[] indexes, object value)
        {
            CheckEntities();

            switch (_defaultItem)
            {
                case DefaultItemSupport.PropertyDefault:
                    InvokePropertySet("_Default", indexes, value);
                    return true;
                case DefaultItemSupport.MethodDefault:
                    InvokeMethod("_Default", indexes, value);
                    return true;
                case DefaultItemSupport.PropertyItem:
                    InvokePropertySet("Item", indexes, value);
                    return true;
                case DefaultItemSupport.MethodItem:
                    InvokeMethod("Item", indexes, value);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Provides the implementation for operations that get member values.
        /// </summary>
        /// <param name="binder">Provides information about the object that called the dynamic operation.</param>
        /// <param name="result">The result of the get operation.</param>
        /// <returns>true if the operation is successful; otherwise, false.</returns>
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            if (IsEnumeratorBinder(binder))
            {
                result = InvokeEnumerator();
                return true;
            }
            else
            {
                result = InvokePropertyGet(binder.Name);
                return true;
            }
        }

        /// <summary>
        /// Provides the implementation for operations that set member values.
        /// </summary>
        /// <param name="binder">Provides information about the object that called the dynamic operation.</param>
        /// <param name="value">The value to set to the member.</param>
        /// <returns> true if the operation is successful; otherwise, false.</returns>
        public override bool TrySetMember(SetMemberBinder binder, object value)
        {

            InvokePropertySet(binder.Name, new object[] { value });
            return true;
        }

        /// <summary>
        /// Provides the implementation for operations that invoke a member.
        /// </summary>
        /// <param name="binder">Provides information about the dynamic operation.</param>
        /// <param name="args">The arguments that are passed to the object member during the invoke operation.</param>
        /// <param name="result">The result of the member invocation.</param>
        /// <returns>true if the operation is successful; otherwise, false.</returns>
        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            if (IsEnumeratorBinder(binder))
            {
                result = InvokeEnumerator();
                return true;
            }
            else
            {
                result = InvokeMethod(binder.Name, args);
                return true;
            }
        }

        #endregion
    }
}
