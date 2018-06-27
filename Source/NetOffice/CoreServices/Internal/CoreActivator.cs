using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using NetOffice.Exceptions;
using NetOffice.Attributes;

namespace NetOffice.CoreServices.Internal
{
    /// <summary>
    /// Core Activation Services
    /// </summary>
    internal class CoreActivator : ICoreActivator
    {
        #region Fields

        private Dictionary<Type, Type> _customTypes = new Dictionary<Type, Type>();

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">affected netoffice core</param>
        /// <exception cref="ArgumentNullException"></exception>
        internal CoreActivator(Core parent)
        {
            if (null == parent)
                throw new ArgumentNullException("parent");
            Parent = parent;
        }

        #endregion

        #region ICOMObjectActivator

        /// <summary>
        /// Occours when a new COMObject instance has been created
        /// </summary>
        public event OnCreateInstanceEventHandler CreateInstance;

        /// <summary>
        /// Occurs when a new COMDynamicObject instance should be created
        /// </summary>
        public event OnCreateCOMDynamicEventHandler CreateDynamicInstance;

        /// <summary>
        /// Occurs when a new COMProxyShare instance should be created
        /// </summary>
        public event OnCreateProxyShareEventHandler CreateProxyShare;

        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        public Core Parent { get; private set; }

        /// <summary>
        /// Registered custom types
        /// </summary>
        public IEnumerable<KeyValuePair<Type, Type>> RegisteredTypes
        {
            get
            {
                return _customTypes.ToArray();
            }
        }

        /// <summary>
        /// Add a custom type
        /// </summary>
        /// <param name="contract">target contract</param>
        /// <param name="implementation">custom implementation</param>
        public void RegisterType(Type contract, Type implementation)
        {
            _customTypes.Add(contract, implementation);
        }

        /// <summary>
        /// Remove a custom type
        /// </summary>
        /// <param name="contract">target contract</param>
        /// <returns>true, if removed, otherwise false</returns>
        public bool UnRegisterType(Type contract)
        {
            return _customTypes.Remove(contract);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Creates a new COMProxyShare instance
        /// </summary>
        /// <param name="sender">requested instance</param>
        /// <param name="comProxy">inner proxy rcw</param>
        ///  <param name="isEnumerator">indicates rcw is an enumerator</param>
        /// <returns>new instance</returns>
        /// <exception cref="CreateCOMProxyShareException">throws when its failed to create instance</exception>
        internal COMProxyShare CreateNewProxyShare(ICOMObject sender, object comProxy, bool isEnumerator)
        {
            try
            {
                COMProxyShare instance = RaiseCreateProxyShare(sender, isEnumerator);
                return null != instance ? instance : new COMProxyShare(Parent, comProxy, isEnumerator);
            }
            catch (Exception exception)
            {
                throw new CreateCOMProxyShareException(exception);
            }
        }

        /// <summary>
        /// Creates a new COMProxyShare instance
        /// </summary>
        /// <param name="sender">requested instance</param>
        /// <param name="comProxy">inner proxy rcw</param>
        /// <returns>new instance</returns>
        /// <exception cref="CreateCOMProxyShareException">throws when its failed to create instance</exception>
        internal COMProxyShare CreateNewProxyShare(ICOMObject sender, object comProxy)
        {
            try
            {
                COMProxyShare instance = RaiseCreateProxyShare(sender, false);
                return null != instance ? instance : new COMProxyShare(Parent, comProxy);
            }
            catch (Exception exception)
            {
                throw new CreateCOMProxyShareException(exception);
            }
        }

        /// <summary>
        /// Try to replace new created instance on CreateInstance event
        /// </summary>
        /// <param name="caller">parent instance</param>
        /// <param name="instance">origin instance</param>
        /// <returns>replace instance or origin instance</returns>
        /// <exception cref="ArgumentNullException">instance is null</exception>
        internal ICOMObject TryReplaceInstance(ICOMObject caller, ICOMObject instance)
        {
            if (null == instance)
                throw new ArgumentNullException("instance");

            ICOMObject result = instance;
            ICOMObject replaceInstance = null;
            RaiseCreateInstance(instance, ref replaceInstance);
            instance.DisposeChildInstances();

            if (null != replaceInstance)
            {
                ProceedReplaceByEventInstance(caller, instance, replaceInstance);
                result = replaceInstance;
            }
            else if (_customTypes.Count > 0)
            {
                Type targetInterface = null;
                var interfaces = instance.InstanceType.GetInterfaces().Where(e => e.GetCustomAttribute<EntityTypeAttribute>() != null);
                int interfacesCount = interfaces.Count();

                if (interfacesCount == 1)
                {
                    targetInterface = interfaces.First();
                }
                else if (interfacesCount > 1)
                {
                    targetInterface = interfaces.FirstOrDefault(e => e.GetCustomAttribute<EntityTypeAttribute>().Type == EntityType.IsCoClass);
                    if (null == targetInterface)
                    {
                        var exceptInheritedInterfaces = interfaces.Except(interfaces.SelectMany(t => t.GetInterfaces()));
                        targetInterface = exceptInheritedInterfaces.FirstOrDefault();
                    }
                }

                Type typeToReplace = null;
                if (null != targetInterface && _customTypes.TryGetValue(targetInterface, out typeToReplace))
                {
                    result = CreateInstanceInternal(typeToReplace, caller, instance, instance.UnderlyingObject, instance.UnderlyingType);
                }
            }

            return result;
        }

        /// <summary>
        /// Raise CreateInstance event
        /// </summary>
        /// <param name="instance">origin instance</param>
        /// <param name="replace">type to replace the instance</param>
        internal void RaiseCreateInstance(ICOMObject instance, ref ICOMObject replace)
        {
            var handler = CreateInstance;
            if (null != handler)
            {
                OnCreateInstanceEventArgs args = new OnCreateInstanceEventArgs(instance);
                handler(Parent, args);
                replace = args.Replace;
            }
        }

        /// <summary>
        /// Raise the CreateDynamicInstance event
        /// </summary>
        /// <param name="instance">requested instance</param>
        /// <param name="comProxy">target proxy</param>
        /// <returns>COMDynamicObject instance or null</returns>
        internal COMDynamicObject RaiseCreateCOMDynamic(ICOMObject instance, object comProxy)
        {
            if (null != CreateDynamicInstance)
            {
                OnCreateCOMDynamicEventArgs args = new OnCreateCOMDynamicEventArgs(instance, comProxy);
                CreateDynamicInstance(Parent, args);
                return args.Replace;
            }
            else
                return null;
        }

        /// <summary>
        /// Initialize replace instance if necessary
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="instance">instance to replace</param>
        /// <param name="replaceInstance">new replaced instance</param>
        private void ProceedReplaceByEventInstance(ICOMObject caller, ICOMObject instance, ICOMObject replaceInstance)
        {
            ICOMObjectInitialize init = replaceInstance as ICOMObjectInitialize;
            if (null != init && false == init.IsInitialized)
            {
                init.InitializeCOMObject(Parent, caller, instance.UnderlyingObject, instance.UnderlyingType);
            }
            if (null != caller)
                caller.RemoveChildObject(instance);
            Parent.InternalObjectRegister.RemoveObjectFromList(instance, null);
        }

        /// <summary>
        /// Raise the CreateProxyShare event
        /// </summary>
        /// <param name="instance">requested instance</param>
        /// <param name="isEnumerator">indicates rcw is an enumerator</param>
        /// <returns>CreateProxyShare instance or null</returns>
        private COMProxyShare RaiseCreateProxyShare(ICOMObject instance, bool isEnumerator)
        {
            var handler = CreateProxyShare;
            if (null != handler)
            {
                OnCreateProxyShareEventArgs args = new OnCreateProxyShareEventArgs(instance, isEnumerator);
                handler(Parent, args);
                return args.Result;
            }
            else
                return null;
        }

        /// <summary>
        /// Creates and initialize an instance without a factory
        /// </summary>
        /// <param name="type">type to create</param>
        /// <param name="caller">caller</param>
        /// <param name="instance">replaced instance</param>
        /// <param name="comProxy">com proxy</param>
        /// <param name="comProxyType">com proxy type</param>
        /// <returns></returns>
        private ICOMObject CreateInstanceInternal(Type type, ICOMObject caller, ICOMObject instance, object comProxy, Type comProxyType)
        {
            ICOMObject result = ComActivator.CreateInitializeInstanceWithoutFactory(instance.Factory, type, caller, comProxy, comProxyType) as ICOMObject;
            if (null != result)
            {
                if(null != caller)
                    caller.RemoveChildObject(instance);
                Parent.InternalObjectRegister.RemoveObjectFromList(instance, null);
            }
            return result;
        }

        #endregion
    }
}
