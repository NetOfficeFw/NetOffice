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
        /// The mapped object
        /// </summary>
        object UnderlyingObject { get; }

        /// <summary>
        /// Name of UnderlyingObject type
        /// </summary>
        string UnderlyingTypeName { get; }

        /// <summary>
        /// Type info of UnderlyingObject
        /// </summary>
        Type InstanceType { get; }

        /// <summary>
        /// The Instance they has been created this instance
        /// </summary>
        COMObject ParentObject { get; set; }
    }
}
