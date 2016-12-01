using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using COMTypes = System.Runtime.InteropServices.ComTypes;
using System.Dynamic;
using System.Collections;

namespace NetOffice
{
    /*
        This is designed to use as dynamic in C# or as object in visual basic.
        This allows to use dynamic late-binding with proxy managed service from Netoffice.(best of both worlds)

        NetOffice.Settings.EnableDynamicObjects(true by default) want enable
        the behavior that Netoffice returns a COMDynamicObject instance if its
        failed to resolve a wrapper class for a com proxy.
    */

    /// <summary>
    /// Represents a managed COM proxy with dynamic runtime type informations.
    /// </summary>
    [DebuggerDisplay("{InstanceFriendlyName}")]
    [TypeConverter(typeof(COMDynamicObjectExpandableObjectConverter))]
    public class COMDynamicObject : DynamicObject, ICOMObject
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
        /// FriendlyTypeName cache field
        /// </summary>
        private string _friendlyTypeName;

        /// <summary>
        /// UnderlyingTypeName cache field
        /// </summary>
        private string _underlyingTypeName;

        /// <summary>
        /// Name of the proxy hosting component
        /// </summary>
        private string _underlyingComponentName;

        /// <summary>
        /// Monitor lock object for accessing the child list
        /// </summary>
        private object _childListLock = new object();

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
            
            Factory = Core.Default;
            ParentObject = null;
            UnderlyingObject = comProxy;
            UnderlyingType = comProxy.GetType();
            Factory.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();
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
            
            if (null != parentObject)
                Factory = parentObject.Factory;
            else
                Factory = Core.Default;

            ParentObject = parentObject;
            UnderlyingObject = comProxy;
            UnderlyingType = comProxy.GetType();
            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                ParentObject.AddChildObject(this);

            Factory.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();
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
            UnderlyingObject = comObject.UnderlyingObject;
            UnderlyingType = comObject.UnderlyingType;

            Factory.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();
        }

        /// <summary>
        /// Creates new instance with given proxy and parent info
        /// </summary>
        /// <param name="factory">current factory instance or null for defauslt</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>       
        public COMDynamicObject(Core factory, ICOMObject parentObject, object comProxy)
        {
            if (null == factory)
                factory = Core.Default;
            Factory = factory;

            ParentObject = parentObject;
            UnderlyingObject = comProxy;
            UnderlyingType = comProxy.GetType();

            if (Settings.Default.EnableProxyManagement && !Object.ReferenceEquals(parentObject, null))
                ParentObject.AddChildObject(this);

            Factory.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();
        }

        /// <summary>
        /// Create new instance from given progid
        /// </summary>
        /// <param name="factory">used factory core</param>
        /// <param name="progId"></param>
        public COMDynamicObject(Core factory, string progId)
        {
            if (String.IsNullOrEmpty(progId))
                throw new ArgumentNullException("progId");

            UnderlyingType = System.Type.GetTypeFromProgID(progId, true);
            UnderlyingObject = Activator.CreateInstance(UnderlyingType);

            Factory = null != factory ? factory : Core.Default;        
            Factory.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();
        }

        /// <summary>
        /// Create new instance from given progid
        /// </summary>
        /// <param name="progId"></param>
        public COMDynamicObject(string progId)
        {
            if (String.IsNullOrEmpty(progId))
                throw new ArgumentNullException("progId");

            UnderlyingType = System.Type.GetTypeFromProgID(progId, true);
            UnderlyingObject = Activator.CreateInstance(UnderlyingType);

            Factory = Core.Default;     
            Factory.AddObjectToList(this);
            _listChildObjects = new List<ICOMObject>();
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
        public static COMDynamicObject ConvertTo(ICOMObject comObject)
        {
            return ConvertTo(comObject, true);
        }

        /// <summary>
        /// Create a COMDynamicObject shallow copy from COMObject instance.
        /// The shallow copy is a root instance in com proxy management without child instances.
        /// Given COMObject instance and shallow copy share the same proxy.
        /// </summary>
        /// <param name="comObject">COMObject instance</param>
        /// <param name="throwException">throw exception if failed to convert</param>
        /// <returns>COMDynamicObject shallow copy</returns>
        public static COMDynamicObject ConvertTo(ICOMObject comObject, bool throwException)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            COMObject instance = comObject as COMObject;
            if (null == instance)
            {
                if (throwException)
                    throw new ArgumentException("COMObject instance required.");
                else
                    return null;
            }

            return new COMDynamicObject(instance);
        }

        /// <summary>
        /// Release com proxy
        /// </summary>
        private void ReleaseCOMProxy(IEnumerable<ICOMObject> ownerPath)
        {
            if (!Object.ReferenceEquals(UnderlyingObject, null))
            {
                ICustomAdapter adapter = UnderlyingObject as ICustomAdapter;
                if (null != adapter)
                {
                    Marshal.ReleaseComObject(adapter.GetUnderlyingObject());
                    Marshal.ReleaseComObject(UnderlyingObject);
                }
                else
                {
                    Marshal.ReleaseComObject(UnderlyingObject);
                }

                Factory.RemoveObjectFromList(this, ownerPath);
                UnderlyingObject = null;
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
        /// <returns>IEnumerable instance</returns>
        private System.Collections.IEnumerable InvokeEnumerator()
        {
            CheckEntities();

            switch (_enumerator)
            {
                case EnumeratorSupport.PropertyEnumerator:
                    return new Enumerator(NetOffice.Utils.GetVariantEnumeratorAsProperty(this));
                case EnumeratorSupport.MethodEnumerator:
                    return new Enumerator(NetOffice.Utils.GetVariantEnumeratorAsMethod(this));
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
            object returnItem = Invoker.MethodReturn(this, name);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
            args = Invoker.ValidateParamsArray(args);
            object returnItem = Invoker.MethodReturn(this, name, args);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
        /// Returns a sequence of all dynamic member names.
        /// </summary>
        /// <returns>a sequence that contains dynamic member names.</returns>
        public override IEnumerable<string> GetDynamicMemberNames()
        {
            CheckEntities();
            if (null == _entities)
                return new string[0];

            int i = 0;
            string[] names = new string[_entities.Length];
            foreach (DynamicObjectEntity item in _entities)
            { 
                names[i] = item.Name;
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
            CheckEntities();

            if (binder.Type == typeof(System.Collections.IEnumerable))
            {
                result = InvokeEnumerator();
                return true;
            }
            else
            {
                // Todo: NetOffice 1.7.5 
                // - check target conversion type is available NetOffice wrapper type and convert into
                result = null;
                return false;
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

        #region ICOMObject

        /// <summary>
        /// Returns the native wrapped proxy
        /// </summary>
        public object UnderlyingObject { get; private set; }

        /// <summary>
        /// Returns Type of native proxy
        /// </summary>
        public Type UnderlyingType { get; private set; }

        /// <summary>
        /// Class name from UnderlyingObject
        /// </summary>
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
        /// Returns friendly name for the instance type
        /// </summary>
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
        /// Name of the hosting NetOffice component
        /// </summary>      
        public string UnderlyingComponentName
        {
            get
            {
                if (null == _underlyingComponentName)
                    _underlyingComponentName =  new UnderlyingTypeNameResolver().GetComponentName(this);
                return _underlyingComponentName;
            }
        }

        /// <summary>
        /// Name of the hosting NetOffice component
        /// </summary>
        public string InstanceComponentName
        {
            get
            {
                return "NetOffice.Core";
            }
        }

        /// <summary>
        /// Friendly instance name of the NetOffice Wrapper class
        /// </summary>
        public string InstanceName
        {
            get
            {
                return InstanceType.FullName;
            }
        }

        /// <summary>
        /// Friendly Name of the NetOffice Wrapper class
        /// </summary>       
        public string InstanceFriendlyName
        {
            get
            {
                return "Dynamic";
            }
        }

        /// <summary>
        /// Type informations from ICOMObject instance
        /// </summary>
        public Type InstanceType
        {
            get
            {
                return _instanceType;
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
        /// Returns parent proxy object
        /// </summary>
        public ICOMObject ParentObject { get; private set; }

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

        /// <summary>
        /// These event was called from Dispose and you can skip the dipose operation here if you want. the event can be helpful for troubleshooting if you dont know why your objects beeing disposed
        /// </summary>
        public event OnDisposeEventHandler OnDispose;

        /// <summary>
        /// Add object to child list
        /// </summary>
        /// <param name="childObject">>target child instance</param>
        public void AddChildObject(ICOMObject childObject)
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
        /// Dispose instance and all child instances
        /// </summary>
        public virtual void Dispose()
        {
            Dispose(true);
        }

        /// <summary>
        /// Dispose instance and all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose proxies with events and one or more event recipients</param>
        public virtual void Dispose(bool disposeEventBinding)
        {
            lock (_disposeLock)
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

                IEnumerable<ICOMObject> ownerPath = null;
                if (Factory.HasProxyRemovedRecipients)
                {
                    ownerPath = Core.GetOwnerPath(this);
                }

                // remove himself from parent childlist
                if ((!Object.ReferenceEquals(ParentObject, null)) && (true == removeFromParent))
                {
                    ParentObject.RemoveChildObject(this);
                    ParentObject = null;
                }

                // call quit automaticly if wanted
                if (_callQuitInDispose && Settings.EnableAutomaticQuit)
                    new QuitCaller().TryCall(Settings, Invoker, this);
                
                // release proxy
                ReleaseCOMProxy(ownerPath);

                // clear supportList reference
                _listSupportedEntities = null;

                _isDisposed = true;
                _isCurrentlyDisposing = false;
            }        
        }

        /// <summary>
        /// Dispose all child instances
        /// </summary>
        public virtual void DisposeChildInstances()
        {
            DisposeChildInstances(true);
        }

        private object _disposeChildLock = new object();

        /// <summary>
        /// Dispose all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose proxies with events and one or more event recipients</param>
        public virtual void DisposeChildInstances(bool disposeEventBinding)
        {
            lock (_disposeChildLock)
            {
                foreach (ICOMObject itemObject in _listChildObjects.ToArray())
                {
                    //COMObjectFaults.RemoveParent(itemObject);
                    itemObject.Dispose(disposeEventBinding);
                }
                _listChildObjects.Clear();
            }         
        }

        /// <summary>
        /// Returns information the proxy provides a method or property with given name at runtime
        /// </summary>
        /// <param name="name">name of the enitity</param>
        /// <returns>true if available, otherwise false</returns>
        public bool EntityIsAvailable(string name)
        {
            return EntityIsAvailable(name, SupportEntityType.Both);
        }

        /// <summary>
        /// Returns information the proxy provides a method or property with given name at runtime
        /// </summary>
        /// <param name="name">name of the enitity</param>
        /// <param name="searchType">limit searching for method or property</param>
        /// <returns>true if available, otherwise false</returns>
        public bool EntityIsAvailable(string name, SupportEntityType searchType)
        {
            return new EntityAvailableResolver().Resolve(Factory, ref _listSupportedEntities, searchType, UnderlyingObject, name);
        }

        /// <summary>
        /// Remove object from child list
        /// </summary>
        /// <param name="childObject">target child instance</param>
        public void RemoveChildObject(ICOMObject childObject)
        {
            bool isLocked = false;
            try
            {
                Monitor.Enter(_childListLock);
                isLocked = true;

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

        /// <summary>
        /// Determines whether two COMObject instances are equal.
        /// </summary>
        /// <param name="objectA"></param>
        /// <param name="objectB"></param>
        /// <returns></returns>
        public static bool operator ==(COMDynamicObject objectA, COMObject objectB)
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
        /// Determines whether two COMObject instances are not equal.
        /// </summary>
        /// <param name="objectA">first instance</param>
        /// <param name="objectB">second instance</param>
        /// <returns>true if equal, otherwise false</returns>
        public static bool operator !=(COMDynamicObject objectA, COMObject objectB)
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
        /// Determines whether two COMObject instances are equal.
        /// </summary>
        /// <param name="objectA"></param>
        /// <param name="objectB"></param>
        /// <returns></returns>
        public static bool operator ==(COMDynamicObject objectA, COMDynamicObject objectB)
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
        public static bool operator ==(COMDynamicObject objectA, object objectB)
        {
            if (!Settings.Default.EnableOperatorOverlads)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return true;
            else if (!Object.ReferenceEquals(objectA, null))
                return objectA.EqualsOnServer(objectB as COMDynamicObject);
            else
                return false;
        }
        
        /// <summary>
        /// Determines whether two COMObject instances are equal.
        /// </summary>
        /// <param name="objectA">first instance</param>
        /// <param name="objectB">second instance</param>
        /// <returns>true if equal, otherwise false</returns>
        public static bool operator ==(object objectA, COMDynamicObject objectB)
        {
            if (!Settings.Default.EnableOperatorOverlads)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return true;
            else if (!Object.ReferenceEquals(objectA, null))
            {
                COMDynamicObject a = (objectA as COMDynamicObject);
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
        public static bool operator !=(COMDynamicObject objectA, COMDynamicObject objectB)
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
        public static bool operator !=(COMDynamicObject objectA, object objectB)
        {
            if (!Settings.Default.EnableOperatorOverlads)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return false;
            else if (!Object.ReferenceEquals(objectA, null))
                return !objectA.EqualsOnServer(objectB as COMDynamicObject);
            else
                return true;
        }

        /// <summary>
        /// Determines whether two COMObject instances are not equal.
        /// </summary>
        /// <param name="objectA">first instance</param>
        /// <param name="objectB">second instance</param>
        /// <returns>true if equal, otherwise false</returns>
        public static bool operator !=(object objectA, COMDynamicObject objectB)
        {
            if (!Settings.Default.EnableOperatorOverlads)
                return Object.ReferenceEquals(objectA, objectB);

            if (Object.ReferenceEquals(objectA, null) && Object.ReferenceEquals(objectB, null))
                return false;
            else if (!Object.ReferenceEquals(objectA, null))
            {
                COMDynamicObject a = objectA as COMDynamicObject;
                if (null != a)
                    return !(objectA as COMDynamicObject).EqualsOnServer(objectB);
                else
                    return null == objectB ? false : true;
            }
            else
                return true;
        }

        #endregion
    }
}
