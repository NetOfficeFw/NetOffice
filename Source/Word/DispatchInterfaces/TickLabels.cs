using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface TickLabels 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838992.aspx </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface TickLabels : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822676.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191842.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ChartFont Font { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193955.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838358.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		string NumberFormat { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837706.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		bool NumberFormatLinked { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193742.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		object NumberFormatLocal { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838903.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Enums.XlTickLabelOrientation Orientation { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194703.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 ReadingOrder { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object AutoScaleFont { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838683.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Depth { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840401.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Offset { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197220.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Alignment { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193071.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		bool MultiLevel { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821425.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ChartFormat Format { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837948.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Application { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822196.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Creator { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837650.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		object Delete();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197209.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		object Select();

		#endregion
	}
}
