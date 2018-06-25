using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi
{
    /// <summary>
    /// InlineShapes
    /// </summary>
    [SyntaxBypass]
    public interface InlineShapes_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PublisherApi.ShapeRange get_Range(object index);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Alias for get_Range
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Publisher", 14, 15, 16), Redirect("get_Range")]
        NetOffice.PublisherApi.ShapeRange Range(object index);

        #endregion
    }

    /// <summary>
    /// DispatchInterface InlineShapes 
    /// SupportByVersion Publisher, 14,15,16
    /// </summary>
    [SupportByVersion("Publisher", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("98091C49-9841-4D1A-BE2D-8FFD9C7702CC")]
    public interface InlineShapes : InlineShapes_, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Shape>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        NetOffice.PublisherApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        new NetOffice.PublisherApi.ShapeRange Range { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="var">object var</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.PublisherApi.Shape this[object var] { get; }

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Shape>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        new IEnumerator<NetOffice.PublisherApi.Shape> GetEnumerator();

        #endregion
    }
}
