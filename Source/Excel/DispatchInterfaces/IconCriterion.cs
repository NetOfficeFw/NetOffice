using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface IconCriterion 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196404.aspx </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface IconCriterion : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835010.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Int32 Index { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838415.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlConditionValueTypes Type { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821367.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		object Value { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835585.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Int32 Operator { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838436.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Enums.XlIcon Icon { get; set; }

		#endregion

	}
}
