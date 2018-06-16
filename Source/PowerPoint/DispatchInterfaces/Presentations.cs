using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Presentations
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743968.aspx </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("91493462-5A91-11CF-8700-00AA0060263B")]
	public interface Presentations : Collection
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744974.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744972.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PowerPointApi.Presentation this[object index] { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745733.aspx </remarks>
		/// <param name="withWindow">optional NetOffice.OfficeApi.Enums.MsoTriState WithWindow = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Add(object withWindow);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745733.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Add();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746171.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		/// <param name="withWindow">optional NetOffice.OfficeApi.Enums.MsoTriState WithWindow = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Open(string fileName, object readOnly, object untitled, object withWindow);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746171.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Open(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746171.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Open(string fileName, object readOnly);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746171.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Open(string fileName, object readOnly, object untitled);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		/// <param name="withWindow">optional NetOffice.OfficeApi.Enums.MsoTriState WithWindow = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation OpenOld(string fileName, object readOnly, object untitled, object withWindow);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation OpenOld(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation OpenOld(string fileName, object readOnly);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation OpenOld(string fileName, object readOnly, object untitled);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746209.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void CheckOut(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745034.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		bool CanCheckOut(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744741.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		/// <param name="withWindow">optional NetOffice.OfficeApi.Enums.MsoTriState WithWindow = -1</param>
		/// <param name="openAndRepair">optional NetOffice.OfficeApi.Enums.MsoTriState OpenAndRepair = 0</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Open2007(string fileName, object readOnly, object untitled, object withWindow, object openAndRepair);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744741.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Open2007(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744741.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Open2007(string fileName, object readOnly);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744741.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Open2007(string fileName, object readOnly, object untitled);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744741.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="readOnly">optional NetOffice.OfficeApi.Enums.MsoTriState ReadOnly = 0</param>
		/// <param name="untitled">optional NetOffice.OfficeApi.Enums.MsoTriState Untitled = 0</param>
		/// <param name="withWindow">optional NetOffice.OfficeApi.Enums.MsoTriState WithWindow = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Presentation Open2007(string fileName, object readOnly, object untitled, object withWindow);

		#endregion
	}
}
