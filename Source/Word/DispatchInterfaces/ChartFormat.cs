using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface ChartFormat 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196903.aspx </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface ChartFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845227.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.FillFormat Fill { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195286.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.GlowFormat Glow { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844794.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.LineFormat Line { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822159.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821987.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.PictureFormat PictureFormat { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837873.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ShadowFormat Shadow { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194275.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.SoftEdgeFormat SoftEdge { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196935.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.OfficeApi.TextFrame2 TextFrame2 { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837722.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ThreeDFormat ThreeD { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191935.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Application { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835474.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227331.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Adjustments Adjustments { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230645.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.OfficeApi.Enums.MsoAutoShapeType AutoShapeType { get; set; }

		#endregion

	}
}
