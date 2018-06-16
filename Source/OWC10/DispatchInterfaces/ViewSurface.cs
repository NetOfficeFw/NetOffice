using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ViewSurface 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("EE658610-D8B3-11D2-8F30-00600893B533")]
	public interface ViewSurface : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 hDC { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 hDCInfo { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="color">Int32 color</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_AlphaBlend(Int32 color);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_AlphaBlend
		/// </summary>
		/// <param name="color">Int32 color</param>
		[SupportByVersion("OWC10", 1), Redirect("get_AlphaBlend")]
		Int32 AlphaBlend(Int32 color);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat textFormat</param>
		/// <param name="text">object text</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_TextHeight(NetOffice.OWC10Api.TextFormat textFormat, object text);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_TextHeight
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat textFormat</param>
		/// <param name="text">object text</param>
		[SupportByVersion("OWC10", 1), Redirect("get_TextHeight")]
		Int32 TextHeight(NetOffice.OWC10Api.TextFormat textFormat, object text);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat textFormat</param>
		/// <param name="text">object text</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_TextWidth(NetOffice.OWC10Api.TextFormat textFormat, object text);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_TextWidth
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat textFormat</param>
		/// <param name="text">object text</param>
		[SupportByVersion("OWC10", 1), Redirect("get_TextWidth")]
		Int32 TextWidth(NetOffice.OWC10Api.TextFormat textFormat, object text);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="picture">stdole.Picture picture</param>
		/// <param name="mask">stdole.Picture mask</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false), NativeResult]
		stdole.Picture get_PictureAlphaBlended(stdole.Picture picture, stdole.Picture mask);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_PictureAlphaBlended
		/// </summary>
		/// <param name="picture">stdole.Picture picture</param>
		/// <param name="mask">stdole.Picture mask</param>
		[SupportByVersion("OWC10", 1), Redirect("get_PictureAlphaBlended")]
		stdole.Picture PictureAlphaBlended(stdole.Picture picture, stdole.Picture mask);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="x">Int32 x</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_ScaleX(Int32 x);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_ScaleX
		/// </summary>
		/// <param name="x">Int32 x</param>
		[SupportByVersion("OWC10", 1), Redirect("get_ScaleX")]
		Int32 ScaleX(Int32 x);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_ScaleY(Int32 y);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_ScaleY
		/// </summary>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("OWC10", 1), Redirect("get_ScaleY")]
		Int32 ScaleY(Int32 y);

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		/// <param name="color">Int32 color</param>
		[SupportByVersion("OWC10", 1)]
		void Rectangle(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, Int32 color);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="x1">Int32 x1</param>
		/// <param name="y1">Int32 y1</param>
		/// <param name="x2">Int32 x2</param>
		/// <param name="y2">Int32 y2</param>
		/// <param name="color">Int32 color</param>
		[SupportByVersion("OWC10", 1)]
		void Line(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 x1, Int32 y1, Int32 x2, Int32 y2, Int32 color);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat textFormat</param>
		/// <param name="text">object text</param>
		[SupportByVersion("OWC10", 1)]
		void Text(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, NetOffice.OWC10Api.TextFormat textFormat, object text);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		/// <param name="picture">stdole.Picture picture</param>
		[SupportByVersion("OWC10", 1)]
		void Picture(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, stdole.Picture picture);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		/// <param name="picture">stdole.Picture picture</param>
		/// <param name="mask">stdole.Picture mask</param>
		[SupportByVersion("OWC10", 1)]
		void PictureMasked(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, stdole.Picture picture, stdole.Picture mask);

		#endregion
	}
}
