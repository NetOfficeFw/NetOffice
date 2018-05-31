using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface ProtectedViewWindow 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822367.aspx </remarks>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface ProtectedViewWindow : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		string _Default { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841246.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		string Caption { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194002.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		bool EnableResize { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196282.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		Double Height { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196899.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		Double Left { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835560.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		Double Top { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837823.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		Double Width { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838633.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		bool Visible { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840283.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		string SourceName { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837810.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		string SourcePath { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837081.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Enums.XlProtectedViewWindowState WindowState { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196579.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Workbook Workbook { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194857.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		void Activate();

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197025.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		bool Close();

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838457.aspx </remarks>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Workbook Edit(object writeResPassword, object updateLinks);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838457.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Workbook Edit();

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838457.aspx </remarks>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Workbook Edit(object writeResPassword);

		#endregion
	}
}
