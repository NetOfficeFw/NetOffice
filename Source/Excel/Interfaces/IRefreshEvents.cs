using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IRefreshEvents 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
	public interface IRefreshEvents : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 BeforeRefresh(bool cancel);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="success">bool success</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 AfterRefresh(bool success);

		#endregion
	}
}
