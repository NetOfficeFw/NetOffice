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
	/// DispatchInterface _Components
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "VBIDE", 12, 14, 5.3), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("0002E161-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VBIDEApi.Components))]
    public interface _Components : ICOMObject, IEnumerableProvider<NetOffice.VBIDEApi.Component>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [BaseResult]
        NetOffice.VBIDEApi.Application Application { get; }

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
        NetOffice.VBIDEApi.Component this[object index] { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="component">NetOffice.VBIDEApi.Component component</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void Remove(NetOffice.VBIDEApi.Component component);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="componentType">NetOffice.VBIDEApi.Enums.vbext_ComponentType componentType</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.Component Add(NetOffice.VBIDEApi.Enums.vbext_ComponentType componentType);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.Component Import(string fileName);

        #endregion

        #region IEnumerable<NetOffice.VBIDEApi.Component>

        /// <summary>
        /// SupportByVersion VBIDE, 12,14,5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        new IEnumerator<NetOffice.VBIDEApi.Component> GetEnumerator();

        #endregion
    }
}
