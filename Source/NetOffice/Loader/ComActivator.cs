using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Exceptions;

namespace NetOffice
{
    /// <summary>
    /// Encapsulate Runtime Activator Services
    /// </summary>
    public static class ComActivator
    {
        /// <summary>
        /// Creates new instance and initialize new instance trough ICOMObjectInitialize interface
        /// </summary>
        /// <param name="factory">affected netoffice core</param>
        /// <param name="type">type to create</param>
        /// <param name="typeFactory">factory to create instance from</param>
        /// <param name="parentObject">parent caller</param>
        /// <param name="comProxy">underlying proxy</param>
        /// <param name="comProxyType">underlying proxy type</param>
        /// <returns>newly created instance</returns>
        /// <exception cref="ActivationException">failed to activate or initialize the instance</exception>
        public static ICOMObject CreateInitializeInstance(Core factory, Type type, ITypeFactory typeFactory, ICOMObject parentObject, object comProxy, Type comProxyType)
        {
            try
            {
                var newInstance = typeFactory.CreateInstance(type);
                ICOMObjectInitialize init = (ICOMObjectInitialize)newInstance;
                init.InitializeCOMObject(factory, parentObject, comProxy, comProxyType);
                return newInstance;
            }
            catch (Exception exception)
            {
                throw new ActivationException(exception);
            }
        }

        /// <summary>
        /// Creates new instance and initialize new instance trough ICOMObjectInitialize interface
        /// </summary>
        /// <param name="factory">affected netoffice core</param>
        /// <param name="type">type to create</param>
        /// <param name="parentObject">parent caller</param>
        /// <param name="comProxy">underlying proxy</param>
        /// <param name="comProxyType">underlying proxy type</param>
        /// <returns>newly created instance</returns>
        /// <exception cref="ActivationException">failed to activate or initialize the instance</exception>
        public static ICOMObject CreateInitializeInstanceWithoutFactory(Core factory, Type type, ICOMObject parentObject, object comProxy, Type comProxyType)
        {
            try
            {
                var newInstance = (ICOMObject)Activator.CreateInstance(type);
                ICOMObjectInitialize init = (ICOMObjectInitialize)newInstance;
                init.InitializeCOMObject(factory, parentObject, comProxy, comProxyType);
                return newInstance;
            }
            catch (Exception exception)
            {
                throw new ActivationException(exception);
            }
        }
    }
}
