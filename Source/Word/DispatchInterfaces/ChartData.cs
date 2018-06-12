using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface ChartData 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193319.aspx </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface ChartData : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192338.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Workbook { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838681.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		bool IsLinked { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845236.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		void Activate();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191845.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		void BreakLink();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232086.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		void ActivateChartDataWindow();

		#endregion
	}
}
