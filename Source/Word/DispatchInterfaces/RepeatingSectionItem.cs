using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface RepeatingSectionItem 
	/// SupportByVersion Word, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229924.aspx </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface RepeatingSectionItem : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230032.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231797.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232367.aspx </remarks>
		[SupportByVersion("Word", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232158.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Range Range { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231102.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.RepeatingSectionItem InsertItemBefore();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231720.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.RepeatingSectionItem InsertItemAfter();

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230396.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		void Delete();

		#endregion
	}
}
