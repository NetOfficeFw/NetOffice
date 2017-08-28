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
 	public class ChChartDraw : COMObject
	{
		#pragma warning disable

		#region Type Information

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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ChChartDraw(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ChChartDraw(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartDraw(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartDraw(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartDraw(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartDraw(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartDraw() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartDraw(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChInterior Interior
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChInterior>(this, "Interior", NetOffice.OWC10Api.ChInterior.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChBorder Border
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChBorder>(this, "Border", NetOffice.OWC10Api.ChBorder.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChFont Font
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChFont>(this, "Font", NetOffice.OWC10Api.ChFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChLine Line
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChLine>(this, "Line", NetOffice.OWC10Api.ChLine.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartDrawModesEnum DrawType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartDrawModesEnum>(this, "DrawType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 hDC
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "hDC");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="id">Int32 id</param>
		[SupportByVersion("OWC10", 1)]
		public void BeginObject(Int32 id)
		{
			 Factory.ExecuteMethod(this, "BeginObject", id);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void EndObject()
		{
			 Factory.ExecuteMethod(this, "EndObject");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x0">Int32 x0</param>
		/// <param name="y0">Int32 y0</param>
		/// <param name="x1">Int32 x1</param>
		/// <param name="y1">Int32 y1</param>
		[SupportByVersion("OWC10", 1)]
		public void DrawLine(Int32 x0, Int32 y0, Int32 x1, Int32 y1)
		{
			 Factory.ExecuteMethod(this, "DrawLine", x0, y0, x1, y1);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="right">Int32 right</param>
		/// <param name="bottom">Int32 bottom</param>
		[SupportByVersion("OWC10", 1)]
		public void DrawRectangle(Int32 left, Int32 top, Int32 right, Int32 bottom)
		{
			 Factory.ExecuteMethod(this, "DrawRectangle", left, top, right, bottom);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="right">Int32 right</param>
		/// <param name="bottom">Int32 bottom</param>
		[SupportByVersion("OWC10", 1)]
		public void DrawEllipse(Int32 left, Int32 top, Int32 right, Int32 bottom)
		{
			 Factory.ExecuteMethod(this, "DrawEllipse", left, top, right, bottom);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bstrText">string bstrText</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		[SupportByVersion("OWC10", 1)]
		public void DrawText(string bstrText, Int32 left, Int32 top)
		{
			 Factory.ExecuteMethod(this, "DrawText", bstrText, left, top);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="xValues">object xValues</param>
		/// <param name="yValues">object yValues</param>
		[SupportByVersion("OWC10", 1)]
		public void DrawPolyLine(object xValues, object yValues)
		{
			 Factory.ExecuteMethod(this, "DrawPolyLine", xValues, yValues);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="xValues">object xValues</param>
		/// <param name="yValues">object yValues</param>
		[SupportByVersion("OWC10", 1)]
		public void DrawPolygon(object xValues, object yValues)
		{
			 Factory.ExecuteMethod(this, "DrawPolygon", xValues, yValues);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="text">string text</param>
		[SupportByVersion("OWC10", 1)]
		public object TextWidth(string text)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextWidth", text);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="text">string text</param>
		[SupportByVersion("OWC10", 1)]
		public object TextHeight(string text)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextHeight", text);
		}

		#endregion

		#pragma warning restore
	}
}
