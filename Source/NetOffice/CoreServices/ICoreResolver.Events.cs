using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Contribution;

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
        /// <param name="contractType">netoffice contract type, can be null</param>
        /// <param name="comProxy">native proxy type</param>
        public ResolveEventArgs(ICOMObject caller, Type contractType, object comProxy)
        {
            Caller = caller;
            Contract = contractType;
            Proxy = ProxyInformation.Create(comProxy);
        }

        /// <summary>
        /// Calling instance or null(Nothing in Visual Basic)
        /// </summary>
        public ICOMObject Caller { get; private set; }

        /// <summary>
        /// Type of the target contract. 
        /// Can be null(Nothing in Visual Basic) if its failed to resolve the corresponding factory
        /// </summary>
        public Type Contract { get; private set; }

        /// <summary>
        /// Detailed proxy informations
        /// </summary>
        public ProxyInformation Proxy { get; private set; }

        /// <summary>
        /// Result Instance (NetOffice is going initialize the instance, after complete the resolve step)
        /// </summary>
        public ICOMObject Result { get; set; }
    }

    /// <summary>
    /// Resolve event handler
    /// </summary>
    /// <param name="sender">Core sender instance</param>
    /// <param name="args">args as provided</param>
    public delegate void ResolveEventHandler(ICoreResolver sender, ResolveEventArgs args);

    #endregion
}
