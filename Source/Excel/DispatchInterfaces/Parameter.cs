using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Parameter 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195700.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface Parameter : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835602.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196867.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195219.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838837.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlParameterDataType DataType { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193923.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlParameterType Type { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840215.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		string PromptString { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840790.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		object Value { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195215.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Range SourceRange { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841105.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821654.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		bool RefreshOnChange { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837843.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlParameterType type</param>
		/// <param name="value">object value</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void SetParam(NetOffice.ExcelApi.Enums.XlParameterType type, object value);

		#endregion
	}
}
