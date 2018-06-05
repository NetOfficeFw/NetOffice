using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices
{
    #region Resolve

    /// <summary>
    /// Arguments in Resolve Event
    /// </summary>
    public class ResolveEventArgs
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="fullClassName">target NetOffice class</param>
        /// <param name="comProxy">native proxy type</param>
        public ResolveEventArgs(ICOMObject caller, string fullClassName, Type comProxy)
        {
            Caller = caller;
            FullClassName = fullClassName;
            ComProxy = ComProxy;
        }

        /// <summary>
        /// Calling instance or null(Nothing in Visual Basic)
        /// </summary>
        public ICOMObject Caller { get; private set; }

        /// <summary>
        /// Target NetOffice class as full qualified name
        /// </summary>
        public string FullClassName { get; private set; }

        /// <summary>
        /// Native Proxy Type
        /// </summary>
        public Type ComProxy { get; private set; }

        /// <summary>
        /// Wrapper class to create an instance from 
        /// </summary>
        public Type Result { get; set; }
    }

    /// <summary>
    /// Resolve event handler
    /// </summary>
    /// <param name="sender">Core sender instance</param>
    /// <param name="args">args as provided</param>
    public delegate void ResolveEventHandler(ICoreResolver sender, ResolveEventArgs args);

    #endregion
}
