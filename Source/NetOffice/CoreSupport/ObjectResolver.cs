using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.COMObjectResolver;

namespace NetOffice.CoreSupport
{
    internal class ObjectResolver : ICOMObjectResolver
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">affected netoffice core</param>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        internal ObjectResolver(Core parent)
        {
            if (null == parent)
                throw new ArgumentNullException("parent");
            Parent = parent;
        }

        #endregion

        #region ICOMObjectResolver

        /// <summary>
        /// Occurs when its failed to resolve a wrapper for a recieved com proxy.
        /// This event allows to find and set the corresponding wrapper at hand.
        /// Otherwise NetOffice want create a dynamic instance if possible.
        /// </summary>
        public event COMObjectResolver.ResolveEventHandler Resolve;
    
        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        public Core Parent { get; private set; }

        #endregion

        #region Methods

        /// <summary>
        /// Raise Resolve event
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="fullClassName">target NetOffice class</param>
        /// <param name="comProxy">native proxy type</param>
        /// <returns>type to use or null</returns>
        internal Type RaiseResolve(ICOMObject caller, string fullClassName, Type comProxy)
        {
            if (null != Resolve)
            {
                COMObjectResolver.ResolveEventArgs args = new COMObjectResolver.ResolveEventArgs(caller, fullClassName, comProxy);
                Resolve(this, args);
                return args.Result;
            }
            else
                return null;
        }


        #endregion
    }
}
