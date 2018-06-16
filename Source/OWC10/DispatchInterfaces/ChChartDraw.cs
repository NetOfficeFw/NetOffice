using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ChChartDraw 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("278585C3-D74B-4E30-ACEB-77D4777639E6")]
	public interface ChChartDraw : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChInterior Interior { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChBorder Border { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChFont Font { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.ChLine Line { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartDrawModesEnum DrawType { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 hDC { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="id">Int32 id</param>
		[SupportByVersion("OWC10", 1)]
		void BeginObject(Int32 id);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void EndObject();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x0">Int32 x0</param>
		/// <param name="y0">Int32 y0</param>
		/// <param name="x1">Int32 x1</param>
		/// <param name="y1">Int32 y1</param>
		[SupportByVersion("OWC10", 1)]
		void DrawLine(Int32 x0, Int32 y0, Int32 x1, Int32 y1);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="right">Int32 right</param>
		/// <param name="bottom">Int32 bottom</param>
		[SupportByVersion("OWC10", 1)]
		void DrawRectangle(Int32 left, Int32 top, Int32 right, Int32 bottom);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="right">Int32 right</param>
		/// <param name="bottom">Int32 bottom</param>
		[SupportByVersion("OWC10", 1)]
		void DrawEllipse(Int32 left, Int32 top, Int32 right, Int32 bottom);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bstrText">string bstrText</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		[SupportByVersion("OWC10", 1)]
		void DrawText(string bstrText, Int32 left, Int32 top);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="xValues">object xValues</param>
		/// <param name="yValues">object yValues</param>
		[SupportByVersion("OWC10", 1)]
		void DrawPolyLine(object xValues, object yValues);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="xValues">object xValues</param>
		/// <param name="yValues">object yValues</param>
		[SupportByVersion("OWC10", 1)]
		void DrawPolygon(object xValues, object yValues);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="text">string text</param>
		[SupportByVersion("OWC10", 1)]
		object TextWidth(string text);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="text">string text</param>
		[SupportByVersion("OWC10", 1)]
		object TextHeight(string text);

		#endregion
	}
}
