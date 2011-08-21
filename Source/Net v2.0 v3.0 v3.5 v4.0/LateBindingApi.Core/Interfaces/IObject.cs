using System;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.Core
{
    /// <summary>
    /// root interface of all latebinding objects
    /// </summary>
    public interface IObject : IDisposable
    {      
        /// <summary>
        /// mapped com proxy
        /// </summary>
        object UnderlyingObject { get; }

        /// <summary>
        /// name of com proxy class
        /// </summary>
        string UnderlyingTypeName { get; }

        /// <summary>
        /// Type info of UnderlyingObject
        /// </summary>
        Type InstanceType { get; }

        /// <summary>
        /// the Instance they has been created this instance
        /// </summary>
        COMObject ParentObject { get; set; }
    }
}
