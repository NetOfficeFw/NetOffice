using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface _NumberFormat 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("81FDD9FE-6464-4A19-82AB-878823E85A5E")]
    [CoClassSource(typeof(NetOffice.OWC10Api.NumberFormat))]
    public interface _NumberFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string Code { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="value">object value</param>
		/// <param name="count">optional Int32 count</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string get_Format(object value, object count);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Format
		/// </summary>
		/// <param name="value">object value</param>
		/// <param name="count">optional Int32 count</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Format")]
		string Format(object value, object count);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string get_Format(object value);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Format
		/// </summary>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Format")]
		string Format(object value);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_Width(Int32 hDC, object value);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Width
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Width")]
		Int32 Width(Int32 hDC, object value);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_Height(Int32 hDC, object value);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Height
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Height")]
		Int32 Height(Int32 hDC, object value);

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="hDCInfo">Int32 hDCInfo</param>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		/// <param name="horizontalAlignment">Int32 horizontalAlignment</param>
		/// <param name="verticalAlignment">Int32 verticalAlignment</param>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1)]
		void Render(Int32 hDC, Int32 hDCInfo, Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, Int32 horizontalAlignment, Int32 verticalAlignment, object value);

		#endregion
	}
}
