﻿using System;
using System.ComponentModel;

namespace NetOffice
{
    /// <summary>
    /// Represents a COM proxy wrapper with type information and access to the underlying proxy
    /// </summary>
    public interface ICOMObjectProxy
    {    
        /// <summary>
        /// Underlying COM proxy
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        object UnderlyingObject { get; }

        /// <summary>
        /// Type information of the <see cref="UnderlyingObject"/>
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        Type UnderlyingType { get; }

        /// <summary>
        /// Full type name of the <see cref="UnderlyingObject"/>
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        string UnderlyingTypeName { get; }

        /// <summary>
        /// Friendly type name of the <see cref="UnderlyingObject"/>
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        string UnderlyingFriendlyTypeName { get; }

        /// <summary>
        /// Component name of the <see cref="UnderlyingObject"/>
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        string UnderlyingComponentName { get; }
        
        /// <summary>
        /// Full name of the NetOffice Wrapper class
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        string InstanceName { get; }

        /// <summary>
        /// Full friendly name of the NetOffice Wrapper class
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        string InstanceFriendlyName { get; }

        /// <summary>
        /// Name of the hosting NetOffice component
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        string InstanceComponentName { get; }
       
        /// <summary>
        /// Type information from <see cref="ICOMObject"/> instance
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        Type InstanceType { get; }
    }
}
