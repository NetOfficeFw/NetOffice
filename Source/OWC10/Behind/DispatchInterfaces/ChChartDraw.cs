using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface ChChartDraw 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ChChartDraw : COMObject, NetOffice.OWC10Api.ChChartDraw
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OWC10Api.ChChartDraw);
                return _contractType;
            }
        }
        private static Type _contractType;


		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(ChChartDraw);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ChChartDraw() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.ChInterior Interior
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChInterior>(this, "Interior", typeof(NetOffice.OWC10Api.ChInterior));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.ChBorder Border
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChBorder>(this, "Border", typeof(NetOffice.OWC10Api.ChBorder));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.ChFont Font
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChFont>(this, "Font", typeof(NetOffice.OWC10Api.ChFont));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.ChLine Line
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChLine>(this, "Line", typeof(NetOffice.OWC10Api.ChLine));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.ChartDrawModesEnum DrawType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartDrawModesEnum>(this, "DrawType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 hDC
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "hDC");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="id">Int32 id</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void BeginObject(Int32 id)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BeginObject", id);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void EndObject()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EndObject");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x0">Int32 x0</param>
		/// <param name="y0">Int32 y0</param>
		/// <param name="x1">Int32 x1</param>
		/// <param name="y1">Int32 y1</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void DrawLine(Int32 x0, Int32 y0, Int32 x1, Int32 y1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DrawLine", x0, y0, x1, y1);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="right">Int32 right</param>
		/// <param name="bottom">Int32 bottom</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void DrawRectangle(Int32 left, Int32 top, Int32 right, Int32 bottom)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DrawRectangle", left, top, right, bottom);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="right">Int32 right</param>
		/// <param name="bottom">Int32 bottom</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void DrawEllipse(Int32 left, Int32 top, Int32 right, Int32 bottom)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DrawEllipse", left, top, right, bottom);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bstrText">string bstrText</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void DrawText(string bstrText, Int32 left, Int32 top)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DrawText", bstrText, left, top);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="xValues">object xValues</param>
		/// <param name="yValues">object yValues</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void DrawPolyLine(object xValues, object yValues)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DrawPolyLine", xValues, yValues);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="xValues">object xValues</param>
		/// <param name="yValues">object yValues</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void DrawPolygon(object xValues, object yValues)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DrawPolygon", xValues, yValues);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="text">string text</param>
		[SupportByVersion("OWC10", 1)]
		public virtual object TextWidth(string text)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextWidth", text);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="text">string text</param>
		[SupportByVersion("OWC10", 1)]
		public virtual object TextHeight(string text)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextHeight", text);
		}

		#endregion

		#pragma warning restore
	}
}


