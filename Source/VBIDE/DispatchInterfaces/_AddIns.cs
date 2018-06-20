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
    /// DispatchInterface _AddIns
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "VBIDE", 12, 14, 5.3), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("DA936B62-AC8B-11D1-B6E5-00A0C90F2744")]
    [CoClassSource(typeof(NetOffice.VBIDEApi.Addins))]
    public interface _AddIns : ICOMObject, IEnumerableProvider<NetOffice.VBIDEApi.AddIn>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.VBE VBE { get; }

        /// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("VBIDE", 12, 14, 5.3), ProxyResult]
        object Parent { get; }

        /// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12, 14, 5.3)]
        Int32 Count { get; }

        #endregion

        #region Methods

        /// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("VBIDE", 12, 14, 5.3)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.VBIDEApi.AddIn this[object index] { get; }

        /// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		[SupportByVersion("VBIDE", 12, 14, 5.3)]
        void Update();

        #endregion

        #region IEnumerable<NetOffice.VBIDEApi.AddIn>

        /// <summary>
        /// SupportByVersion VBIDE, 12,14,5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        new IEnumerator<NetOffice.VBIDEApi.AddIn> GetEnumerator();

        #endregion
    }
}
