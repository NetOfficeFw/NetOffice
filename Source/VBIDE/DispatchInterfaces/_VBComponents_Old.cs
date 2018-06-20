using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VBIDEApi
{
    /// <summary>
    /// DispatchInterface _VBComponents_Old
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "VBIDE", 12, 14, 5.3), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("0002E162-0000-0000-C000-000000000046")]
    public interface _VBComponents_Old : ICOMObject, IEnumerableProvider<NetOffice.VBIDEApi.VBComponent>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.VBProject Parent { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.VBE VBE { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.VBIDEApi.VBComponent this[object index] { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void Remove(NetOffice.VBIDEApi.VBComponent vBComponent);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="componentType">NetOffice.VBIDEApi.Enums.vbext_ComponentType componentType</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.VBComponent Add(NetOffice.VBIDEApi.Enums.vbext_ComponentType componentType);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.VBComponent Import(string fileName);

        #endregion

        #region IEnumerable<NetOffice.VBIDEApi.VBComponent>

        /// <summary>
        /// SupportByVersion VBIDE, 12,14,5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        new IEnumerator<NetOffice.VBIDEApi.VBComponent> GetEnumerator();

        #endregion
    }
}
