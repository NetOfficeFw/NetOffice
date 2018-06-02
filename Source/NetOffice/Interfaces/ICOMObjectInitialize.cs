using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// ICOMObject Initialization Tasks
    /// </summary>
    public interface ICOMObjectInitialize
    {
        #region Ctor

        /// <summary>
        /// Instance is already initialized
        /// </summary>
        bool IsInitialized { get; }

        /// <summary>
        /// Initialize instance and replace the given replacedObject in proxy management
        /// all created childs from replacedObject are now childs from the new instance
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="replacedObject">the instance you want replace in current NO proxy management</param>
        void InitializeCOMObject(Core factory, ICOMObject replacedObject);

        /// <summary>
        /// Initialize instance and replace the given replacedObject in proxy management
        /// all created childs from replacedObject are now childs from the new instance
        /// </summary>
        /// <param name="replacedObject">the instance you want replace in current NO proxy management</param>
        void InitializeCOMObject(ICOMObject replacedObject);

        /// <summary>
        /// Initialize instance with given proxy
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="comProxy">the now wrapped comProxy root instance</param>
        void InitializeCOMObject(Core factory, object comProxy);

        /// <summary>
        /// Initialize instance with given proxy and parent info
        /// </summary>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        void InitializeCOMObject(ICOMObject parentObject, object comProxy);

        /// <summary>
        /// Initialize instance with given proxy
        /// </summary>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        void InitializeCOMObject(object comProxy);

        /// <summary>
        /// Initialize instance with given proxy and parent info
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="proxyShare">proxy share instead of proxy</param>
        void InitializeCOMObject(Core factory, ICOMObject parentObject, COMProxyShare proxyShare);

        /// <summary>
        /// Initialize instance with given proxy and parent info
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        void InitializeCOMObject(Core factory, ICOMObject parentObject, object comProxy);

        /// <summary>
        /// Initialize instance with given proxy, parent info and info instance is an enumerator
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        ///  <param name="isEnumerator"></param>
        void InitializeCOMObject(Core factory, ICOMObject parentObject, object comProxy, bool isEnumerator);

        /// <summary>
        /// Initialize instance with given proxy, parent info and info instance is an enumerator
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        /// <param name="isEnumerator">instance is an enumerator</param>
        /// <param name="name">custom instance name</param>
        void InitializeCOMObject(Core factory, ICOMObject parentObject, object comProxy, bool isEnumerator, string name);

        /// <summary>
        /// Initialize instance with given proxy, type info and parent info
        /// </summary>
        /// <param name="factory">current factory instance or null for default</param>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        /// <param name="comProxyType">typeinfo from comProy if you have or null</param>
        void InitializeCOMObject(Core factory, ICOMObject parentObject, object comProxy, Type comProxyType);

        /// <summary>
        /// Initialize instance with given proxy, type info and parent info
        /// </summary>
        /// <param name="parentObject">the parent instance where you have these instance from</param>
        /// <param name="comProxy">the now wrapped comProxy instance</param>
        /// <param name="comProxyType">typeinfo from comProy if you have or null</param>
        void InitializeCOMObject(ICOMObject parentObject, object comProxy, Type comProxyType);

        /// <summary>
        /// Initialize instace with progid
        /// </summary>
        /// <param name="factory">current factory instance</param>
        /// <param name="progId">registered ProgID</param>
        void InitializeCOMObject(Core factory, string progId);

        /// <summary>
        /// Initialize instace with progid
        /// </summary>
        /// <param name="progId">registered ProgID</param>
        void InitializeCOMObject(string progId);

        #endregion
    }
}
