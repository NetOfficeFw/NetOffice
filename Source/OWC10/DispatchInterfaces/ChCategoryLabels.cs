using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OWC10Api
{
    /// <summary>
    /// ChCategoryLabels
    /// </summary>
    [SyntaxBypass]
    public interface ChCategoryLabels_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="level">optional Int32 level</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 get_ItemCount(object level);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_ItemCount
        /// </summary>
        /// <param name="level">optional Int32 level</param>
        [SupportByVersion("OWC10", 1), Redirect("get_ItemCount")]
        Int32 ItemCount(object level);

        #endregion
    }

    /// <summary>
    /// DispatchInterface ChCategoryLabels 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("EDF774DC-D540-40F4-99F2-75C83379CAA8")]
    public interface ChCategoryLabels : ChCategoryLabels_, IEnumerableProvider<NetOffice.OWC10Api.ChCategoryLabel>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        Int32 LevelCount { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new Int32 ItemCount { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.ChAxis Parent { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// Custom Indexer
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        NetOffice.OWC10Api.ChCategoryLabel this[object index] { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        /// <param name="level">optional Int32 level</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OWC10Api.ChCategoryLabel this[object index, object level] { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api.PivotResultGroupAxis PivotAxis { get; }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.ChCategoryLabel>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<NetOffice.OWC10Api.ChCategoryLabel> GetEnumerator();

        #endregion
    }
}
