using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface DiagramNodes 
	/// SupportByVersion PowerPoint, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "PowerPoint", 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("914934DA-5A91-11CF-8700-00AA0060263B")]
	public interface DiagramNodes : ICOMObject, IEnumerableProvider<NetOffice.PowerPointApi.DiagramNode>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16), ProxyResult]
		object Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PowerPointApi.DiagramNode this[object index] { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void SelectAll();

        #endregion

        #region IEnumerable<NetOffice.PowerPointApi.DiagramNode>

        /// <summary>
        /// SupportByVersion PowerPoint, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.PowerPointApi.DiagramNode> GetEnumerator();

        #endregion
    }
}
