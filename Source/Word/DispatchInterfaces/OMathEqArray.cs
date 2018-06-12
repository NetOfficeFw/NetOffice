using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface OMathEqArray 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835153.aspx </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface OMathEqArray : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844882.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840628.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197800.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840021.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.OMathArgs E { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845015.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool MaxDist { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839934.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool ObjDist { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191853.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdOMathVertAlignType Align { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821078.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdOMathSpacingRule RowSpacingRule { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845329.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 RowSpacing { get; set; }

		#endregion

	}
}
