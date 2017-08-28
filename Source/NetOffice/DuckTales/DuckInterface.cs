using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.CompilerServices;
using NetOffice.Attributes;

namespace NetOffice.Duck
{
    internal class DuckInterface
    {
        #region Fields

        private bool? _isValidEventClass;

        private KeyValuePair<string, Type>[] _sinks;

        private EventInfo[] _events;

        private MethodInfo[] _methods;

        private MethodInfo[] _methodsWithSyntaxIssue;

        private MethodInfo[] _methodsWithEnumerator;

        private PropertyInfo[] _properties;

        private PropertyInfo[] _propertiesIndexer;

        private bool? _syntaxClassRequired;

        #endregion

        #region Ctor

        internal DuckInterface(Type interfaceType)
        {
            if (null == interfaceType)
                throw new ArgumentNullException("interfaceType");
            if(!interfaceType.IsInterface)
                throw new ArgumentException("Given type isn't an interface.");

            InterfaceType = interfaceType;

            List<Type> interfaces = interfaceType.GetInterfaces().ToList();
        
            if (!interfaces.Any(e => e == typeof(ICOMObject)))
                throw new ArgumentException("Interface must inherit from ICOMObject.");

            List<Type> knownInterfaces = new List<Type>();

            foreach (Type item in interfaces)
            {
                if (IsNetOfficeCoreAssembly(item.Assembly) || item.FullName == "System.IDisposable")
                    knownInterfaces.Add(item);
                if (HasSyntaxInterfaceAttribute(item))
                    _syntaxClassRequired = true;
            }
            foreach (var item in knownInterfaces)
                interfaces.Remove(item);

            Interfaces = interfaces.ToArray();

            Types = new Type[Interfaces.Length +1];
            Types[0] = InterfaceType;
            for (int i = 0; i < Interfaces.Length; i++)
                Types[i+1] = Interfaces[i];
        }

        #endregion

        #region Properties

        public bool IsValidEventClass
        {
            get
            {
                if (null == _isValidEventClass)
                {
                    object[] attributes = InterfaceType.GetCustomAttributes(typeof(EntityTypeAttribute), false);
                    bool isCoClass = attributes.Length == 1 && (attributes[0] as EntityTypeAttribute).Type == EntityType.IsCoClass;
                    attributes = InterfaceType.GetCustomAttributes(typeof(EventSinkAttribute), false);
                    bool isSinkSupported = false;
                    if (attributes.Length == 1 && (attributes[0] as EventSinkAttribute).Sinks != null)
                    {
                        isSinkSupported = true;
                        Type[] sinks = (attributes[0] as EventSinkAttribute).Sinks;
                        foreach (var item in sinks)
                        {
                            if (!item.IsSubclassOf(typeof(SinkHelper)))
                            {
                                isSinkSupported = false;
                                break;
                            }
                        }
                    }

                    bool implementsEventBinding = InterfaceType.GetInterfaces().Any(e => e.FullName == "NetOffice.IEventBinding");
                    _isValidEventClass = isCoClass && isSinkSupported && implementsEventBinding;
                }

                return _isValidEventClass.Value;
            }           
        }

        public KeyValuePair<string, Type>[] EventSinks
        {
            get
            {
                if (null == _sinks)
                {
                    object[] attributes = InterfaceType.GetCustomAttributes(typeof(EventSinkAttribute), false);
                    if (attributes.Length == 1)
                    {
                        Dictionary<string, Type> result = new Dictionary<string, Type>();
                        EventSinkAttribute sink = (EventSinkAttribute)attributes[0];
                        Type[] sinks = sink.Sinks;
                        foreach (Type item in sinks)
                        {
                            FieldInfo staticIdField = item.GetField("Id", BindingFlags.Static | BindingFlags.Public);
                            if (null != staticIdField)
                            {
                                string id = staticIdField.GetValue(null) as string;
                                if (null != id)
                                    result.Add(id, item);
                            }
                        }
                        _sinks = result.ToArray();
                    }
                    else
                        _sinks = new KeyValuePair<string, Type>[0];
                }
                return _sinks;
            }
        }

        public EventInfo[] Events
        {
            get
            {
                if (null == _events)
                {
                    List<EventInfo> result = new List<EventInfo>();         
                    foreach (Type type in Types)
                    {
                        if (IsNetOfficeCoreAssembly(type.Assembly))
                            continue;
                        EventInfo[] events = type.GetEvents();
                        foreach (EventInfo item in events)
                        {
                            if (!ContainsEvent(result, item))
                                result.Add(item);
                        }
                    }

                    _events = result.ToArray();
                }

                return _events;
            }        
        }

        private static bool IsIndexerProperty(PropertyInfo property)
        {
            return property.GetCustomAttributes(typeof(IndexPropertyAttribute), false).Length > 0;
        }

        public EnumeratorAttribute GetEnumeratorAttribute()
        {
            object[] attributes = InterfaceType.GetCustomAttributes(typeof(EnumeratorAttribute), false);
            if (attributes.Length > 0)
                return (attributes[0] as EnumeratorAttribute);
            else
                return null;
        }

        public HasIndexPropertyAttribute GetHasIndexPropertyAttribute()
        {
            object[] attributes = InterfaceType.GetCustomAttributes(typeof(HasIndexPropertyAttribute), false);
            if (attributes.Length > 0)
                return (attributes[0] as HasIndexPropertyAttribute);
            else
                return null;
        }

        public PropertyInfo[] PropertiesIndexer
        {
            get
            {
                if (null == _propertiesIndexer)
                {
                    List<PropertyInfo> result = new List<PropertyInfo>();
                    foreach (Type type in Types)
                    {
                        if (IsNetOfficeCoreAssembly(type.Assembly))
                            continue;
                        PropertyInfo[] properties = type.GetProperties();
                        foreach (PropertyInfo item in properties)
                        {
                            if (!IsIndexerProperty(item))
                                continue;
                            if (!ContainsProperty(result, item))
                                result.Add(item);
                        }
                    }
                    _propertiesIndexer = result.ToArray();
                }
                return _propertiesIndexer;
            }
        }

        public PropertyInfo[] Properties
        {
            get
            {
                if (null == _properties)
                {
                    List<PropertyInfo> result = new List<PropertyInfo>();
                    foreach (Type type in Types)
                    {
                        if (IsNetOfficeCoreAssembly(type.Assembly))
                            continue;
                        PropertyInfo[] properties = type.GetProperties();
                        foreach (PropertyInfo item in properties)
                        {
                            if (IsIndexerProperty(item))
                                continue;
                            if (!ContainsProperty(result, item))
                                result.Add(item);
                        }
                    }

                    return _properties = result.ToArray();
                }
                return _properties;
            }
        }

        public MethodInfo[] MethodsWithEnumerator
        {
            get
            {
                if (null == _methodsWithEnumerator)
                {
                    List<MethodInfo> result = new List<MethodInfo>();
                    foreach (Type type in Types)
                    {
                        if (IsNetOfficeCoreAssembly(type.Assembly))
                            continue;
                        MethodInfo[] methods = type.GetMethods();
                        foreach (MethodInfo item in methods)
                        {
                            if (!MethodIsEnumerator(item))
                                continue;
                            bool isVisible = MethodIsVisible(item);
                            if (!ContainsMethod(result, item) && true == isVisible)
                                result.Add(item);
                        }
                    }

                    _methodsWithEnumerator = result.ToArray();
                }
                return _methodsWithEnumerator;
            }
        }

        public MethodInfo[] Methods
        {
            get
            {
                if (null == _methods)
                {
                    List<MethodInfo> result = new List<MethodInfo>();
                    foreach (Type type in Types)
                    {
                        if (IsNetOfficeCoreAssembly(type.Assembly))
                            continue;
                        MethodInfo[] methods = type.GetMethods();
                        foreach (MethodInfo item in methods)
                        {
                            if (MethodIsEnumerator(item))
                                continue;
                            bool isVisible = MethodIsVisible(item);
                            if (!ContainsMethod(result, item) && true == isVisible)
                                result.Add(item);
                        }
                    }

                    _methods = result.ToArray();
                }
                return _methods;
            }
        }

        public MethodInfo[] MethodsWithSyntaxIssue
        {
            get
            {
                if (null == _methodsWithSyntaxIssue)
                {
                    List<MethodInfo> result = new List<MethodInfo>();
                    foreach (Type type in Types)
                    {
                        if (IsNetOfficeCoreAssembly(type.Assembly))
                            continue;
                        MethodInfo[] methods = type.GetMethods();
                        foreach (MethodInfo item in methods)
                        {
                            if (MethodIsEnumerator(item))
                                continue;
                            bool isVisible = MethodIsSyntaxIssue(item);
                            if (!ContainsMethod(result, item) && true == isVisible)
                                result.Add(item);
                        }
                    }

                    _methodsWithSyntaxIssue = result.ToArray();
                }
                return _methodsWithSyntaxIssue;
            }            
        }

        public string AssemblyName
        {
            get
            {
                return InterfaceType.Assembly.GetName().Name + ".dll";
            }
        }

        public string AssemblyFullName
        {
            get
            {
                return InterfaceType.Assembly.FullName;
            }
        }

        public string FullName
        {
            get
            {
                return InterfaceType.FullName;
            }
        }

        public string Name
        {
            get
            {
                return InterfaceType.Name;
            }
        }

        public bool SyntaxClassRequired
        {
            get
            {
                if (null == _syntaxClassRequired)
                {
                    _syntaxClassRequired = MethodsWithSyntaxIssue.Length > 0;
                }
                return _syntaxClassRequired.Value;
            }
        }

        public Type InterfaceType { get; private set; }

        private Type[] Interfaces { get; set; }

        private Type[] Types { get; set; }

        #endregion

        #region Methods

        private static bool MethodIsEnumerator(MethodInfo method)
        {          
            return method.ReturnType.FullName.StartsWith("System.Collections.Generic.IEnumerator") ||
                method.ReturnType.FullName.StartsWith("System.Collections.IEnumerator");
        }
        
        private static bool MethodIsSyntaxIssue(MethodInfo method)
        {
            return HasSyntaxInterfaceAttribute(method.DeclaringType);
        }

        private static bool MethodIsVisible(MethodInfo method)
        {
            if (HasSyntaxInterfaceAttribute(method.DeclaringType))
                return false;

            else if (method.Name.StartsWith("get_") || method.Name.StartsWith("set_") ||
                method.Name.StartsWith("add_") || method.Name.StartsWith("remove_"))
            {
                object[] attributes = method.GetCustomAttributes(typeof(VisibleAttribute), false);
                if (attributes.Length > 0)
                {
                    VisibleAttribute visible = attributes[0] as VisibleAttribute;
                    return visible.Value;
                }
                else
                    return false;
            }
            else
                return true;
        }

        private static bool HasSyntaxInterfaceAttribute(Type type)
        {
            object[] attributes = type.GetCustomAttributes(typeof(SyntaxBypassAttribute), false);
            return attributes.Length > 0;
        }

        private static bool ContainsProperty(List<PropertyInfo> list, PropertyInfo property)
        {
            foreach (PropertyInfo item in list)
            {
                if (IsPropertyEqual(item, property))
                    return true;
            }
            return false;
        }

        private static bool IsPropertyEqual(PropertyInfo propertyA, PropertyInfo propertyB)
        {
            if (propertyA.Name != propertyB.Name)
                return false;

            ParameterInfo[] indexArgumentsA = propertyA.GetIndexParameters();
            ParameterInfo[] indexArgumentsB = propertyB.GetIndexParameters();

            if (indexArgumentsA.Length != indexArgumentsB.Length)
                return false;

            return true;
        }

        private static bool ContainsMethod(List<MethodInfo> list, MethodInfo method)
        {
            foreach (MethodInfo item in list)
            {
                if (IsMethodEqual(item, method))
                    return true;
            }
            return false;
        }

        private static bool IsMethodEqual(MethodInfo methodA, MethodInfo methodB)
        {
            if (methodA.Name != methodB.Name)
                return false;

            ParameterInfo[] argsA = methodA.GetParameters();
            ParameterInfo[] argsB = methodB.GetParameters();

            if (argsA.Length != argsB.Length)
                return false;

            for (int i = 0; i < argsA.Length; i++)
            {
                ParameterInfo argA = argsA[i];
                ParameterInfo argB = argsB[i];

                if (argA.ParameterType != argB.ParameterType)
                    return false;
            }
            
            return true;
        }

        private static bool ContainsEvent(List<EventInfo> list, EventInfo method)
        {
            foreach (EventInfo item in list)
            {
                if (IsEventEqual(item, method))
                    return true;
            }
            return false;
        }

        private static bool IsEventEqual(EventInfo methodA, EventInfo methodB)
        {
            if (methodA.Name != methodB.Name)
                return false;
          
            return true;
        }

        private static bool IsNetOfficeCoreAssembly(Assembly assembly)
        {
            object[] attributes = assembly.GetCustomAttributes(typeof(System.Runtime.InteropServices.GuidAttribute), false);
            if (attributes.Length == 1)
            {
                System.Runtime.InteropServices.GuidAttribute attribute = attributes[0] as System.Runtime.InteropServices.GuidAttribute;
                return attribute.Value.Equals("ac0714f2-3d04-11d1-ae7d-00a0c90f26f4", StringComparison.InvariantCultureIgnoreCase);
            }
            return false;
        }
       
        #endregion
    }
}