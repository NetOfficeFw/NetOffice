using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using NetOffice.COMObjectActivator;
using NetOffice.Exceptions;

namespace NetOffice.CoreSupport
{
    internal class ObjectActivator : ICOMObjectActivator
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">affected netoffice core</param>
        /// <exception cref="ArgumentNullException"></exception>
        internal ObjectActivator(Core parent)
        {
            if (null == parent)
                throw new ArgumentNullException("parent");
            Parent = parent;
        }

        #endregion

        #region Parent

        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        public Core Parent { get; private set; }

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
        /// <param name="comProxyType">type of native com proxy</param>
        /// <returns>replace instance or origin instance</returns>
        internal ICOMObject TryReplaceInstance(ICOMObject caller, ICOMObject instance, Type comProxyType)
        {
            ICOMObject result = instance;
            Type typeToReplace = null;
            RaiseCreateInstance(instance, ref typeToReplace);
            instance.DisposeChildInstances();

            if (null != typeToReplace)
            {
                ICOMObject replaceInstance = ComActivator.CreateInitializeInstance(typeToReplace, caller, instance.UnderlyingObject, comProxyType) as ICOMObject;
                if (null != replaceInstance)
                {
                    caller.RemoveChildObject(instance);
                    Parent.ObjectList.RemoveObjectFromList(instance, null);
                    result = replaceInstance;
                }
            }

            return result;
        }

        /// <summary>
        /// Raise CreateInstance event
        /// </summary>
        /// <param name="instance">origin instance</param>
        /// <param name="replace">type to replace the instance</param>
        internal void RaiseCreateInstance(ICOMObject instance, ref Type replace)
        {
            var handler = CreateInstance;
            if (null != handler)
            {
                OnCreateInstanceEventArgs args = new OnCreateInstanceEventArgs(instance);
                handler(this, args);
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
                CreateDynamicInstance(this, args);
                return args.Result;
            }
            else
                return null;
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
                handler(this, args);
                return args.Result;
            }
            else
                return null;
        }

        #endregion
    }
}
