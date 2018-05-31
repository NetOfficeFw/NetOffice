using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface ModelTable 
	/// SupportByVersion Excel, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231552.aspx </remarks>
	[SupportByVersion("Excel", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface ModelTable : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231473.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227692.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230584.aspx </remarks>
		[SupportByVersion("Excel", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229029.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230844.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		string SourceName { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230801.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.ModelTableColumns ModelTableColumns { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231266.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.WorkbookConnection SourceWorkbookConnection { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		DateTime LastRefresh { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230118.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		Int32 RecordCount { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228459.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		void Refresh();

		#endregion
	}
}
