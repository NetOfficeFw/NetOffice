using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface ChartCategory 
	/// SupportByVersion Word, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227634.aspx </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface ChartCategory : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231175.aspx </remarks>
		[SupportByVersion("Word", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232152.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227534.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		bool IsFiltered { get; set; }

		#endregion

	}
}
