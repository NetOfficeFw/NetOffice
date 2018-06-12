using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface LegendKey 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845167.aspx </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface LegendKey : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195907.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.WordApi.ChartBorder Border { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.WordApi.Interior Interior { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.WordApi.ChartFillFormat Fill { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196674.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		bool InvertIfNegative { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840123.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 MarkerBackgroundColor { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193726.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Enums.XlColorIndex MarkerBackgroundColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822183.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 MarkerForegroundColor { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839802.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Enums.XlColorIndex MarkerForegroundColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197567.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 MarkerSize { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836606.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Enums.XlMarkerStyle MarkerStyle { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821020.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 PictureType { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Double PictureUnit { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837952.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		bool Smooth { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197923.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Double Left { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821304.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Double Top { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193958.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Double Width { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845574.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Double Height { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195185.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		bool Shadow { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822640.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ChartFormat Format { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196297.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Application { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839984.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195506.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Double PictureUnit2 { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194675.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		object ClearFormats();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839327.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		object Delete();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		object Select();

		#endregion
	}
}
