using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Printer 
	/// SupportByVersion Access, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("DBC5175F-A8ED-11D3-A0DD-00C04F68712B")]
    [CoClassSource(typeof(NetOffice.AccessApi.Printer))]
    public interface _Printer : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194552.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		NetOffice.AccessApi.Enums.AcPrintColor ColorMode { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193471.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 Copies { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822789.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		string DeviceName { get; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195870.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		string DriverName { get; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198051.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		NetOffice.AccessApi.Enums.AcPrintDuplex Duplex { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191910.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		NetOffice.AccessApi.Enums.AcPrintOrientation Orientation { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834798.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		NetOffice.AccessApi.Enums.AcPrintPaperBin PaperBin { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836635.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		NetOffice.AccessApi.Enums.AcPrintPaperSize PaperSize { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845317.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		string Port { get; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195844.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		NetOffice.AccessApi.Enums.AcPrintObjQuality PrintQuality { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194827.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 LeftMargin { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834469.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 RightMargin { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835658.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 TopMargin { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835336.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 BottomMargin { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192121.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		bool DataOnly { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195704.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 ItemsAcross { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196146.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 RowSpacing { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844923.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 ColumnSpacing { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822094.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		bool DefaultSize { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196498.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 ItemSizeWidth { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196765.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		Int32 ItemSizeHeight { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194662.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		NetOffice.AccessApi.Enums.AcPrintItemLayout ItemLayout { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		bool IsMemberSafe(Int32 dispid);

		#endregion
	}
}
