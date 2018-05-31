using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface HeaderFooter 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823128.aspx </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface HeaderFooter : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197191.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		string Text { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820839.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Graphic Picture { get; }

		#endregion

	}
}
