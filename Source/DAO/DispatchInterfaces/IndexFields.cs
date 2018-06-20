using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.DAOApi
{
    /// <summary>
    /// DispatchInterface IndexFields 
    /// SupportByVersion DAO, 3.6,12.0
    /// </summary>
    [SupportByVersion("DAO", 3.6, 12.0)]
    [EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("0000005D-0000-0010-8000-00AA006D2EA4")]
    public interface IndexFields : _DynaCollection
    {
        #region Properties

        /// <summary>
        /// SupportByVersion DAO 3.6, 12.0
        /// Get
        /// </summary>
        /// <param name="item">optional object item</param>
        [SupportByVersion("DAO", 3.6, 12.0)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        object this[object item] { get; }

        #endregion
    }
}
