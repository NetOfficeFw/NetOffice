using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface FreeformBuilder 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00021266-0000-0000-C000-000000000046")]
	public interface FreeformBuilder : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		/// <param name="x2">optional object x2</param>
		/// <param name="y2">optional object y2</param>
		/// <param name="x3">optional object x3</param>
		/// <param name="y3">optional object y3</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void AddNodes(NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1, object x2, object y2, object x3, object y3);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void AddNodes(NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		/// <param name="x2">optional object x2</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void AddNodes(NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1, object x2);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		/// <param name="x2">optional object x2</param>
		/// <param name="y2">optional object y2</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void AddNodes(NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1, object x2, object y2);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		/// <param name="x2">optional object x2</param>
		/// <param name="y2">optional object y2</param>
		/// <param name="x3">optional object x3</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void AddNodes(NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1, object x2, object y2, object x3);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape ConvertToShape();

		#endregion
	}
}
