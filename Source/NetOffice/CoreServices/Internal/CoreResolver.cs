using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices.Internal
{
    /// <summary>
    /// Core Type Resolver
    /// </summary>
    internal class CoreResolver : ICoreResolver
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">affected netoffice core</param>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        internal CoreResolver(Core parent)
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
        public event ResolveEventHandler Resolve;
    
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
        /// <param name="contractType">netoffice contract type, can be null</param>
        /// <param name="comProxy">native proxy</param>
        /// <returns>type to use or null</returns>
        internal ICOMObject RaiseResolve(ICOMObject caller, Type contractType, object comProxy)
        {
            var handler = Resolve;
            if (null != handler)
            {
                ResolveEventArgs args = new ResolveEventArgs(caller, contractType, comProxy);
                handler(Parent, args);
                return args.Result;
            }
            else
                return null;
        }

        #endregion
    }
}
