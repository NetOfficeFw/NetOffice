using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface ConnectorFormat 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820940.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface ConnectorFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838840.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838381.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839785.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196008.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState BeginConnected { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837084.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Shape BeginConnectedShape { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821637.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 BeginConnectionSite { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840089.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState EndConnected { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840101.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Shape EndConnectedShape { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820998.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 EndConnectionSite { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820905.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoConnectorType Type { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840831.aspx </remarks>
		/// <param name="connectedShape">NetOffice.ExcelApi.Shape connectedShape</param>
		/// <param name="connectionSite">Int32 connectionSite</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void BeginConnect(NetOffice.ExcelApi.Shape connectedShape, Int32 connectionSite);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194927.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void BeginDisconnect();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839018.aspx </remarks>
		/// <param name="connectedShape">NetOffice.ExcelApi.Shape connectedShape</param>
		/// <param name="connectionSite">Int32 connectionSite</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void EndConnect(NetOffice.ExcelApi.Shape connectedShape, Int32 connectionSite);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198362.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void EndDisconnect();

		#endregion
	}
}
