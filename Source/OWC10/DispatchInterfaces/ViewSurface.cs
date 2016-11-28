using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// DispatchInterface ViewSurface 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ViewSurface : COMObject
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(ViewSurface);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ViewSurface(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewSurface(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewSurface(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewSurface(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewSurface(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewSurface() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewSurface(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 hDC
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "hDC", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 hDCInfo
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "hDCInfo", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="color">Int32 Color</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_AlphaBlend(Int32 color)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(color);
			object returnItem = Invoker.PropertyGet(this, "AlphaBlend", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_AlphaBlend
		/// </summary>
		/// <param name="color">Int32 Color</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 AlphaBlend(Int32 color)
		{
			return get_AlphaBlend(color);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat TextFormat</param>
		/// <param name="text">object Text</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_TextHeight(NetOffice.OWC10Api.TextFormat textFormat, object text)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(textFormat, text);
			object returnItem = Invoker.PropertyGet(this, "TextHeight", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_TextHeight
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat TextFormat</param>
		/// <param name="text">object Text</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 TextHeight(NetOffice.OWC10Api.TextFormat textFormat, object text)
		{
			return get_TextHeight(textFormat, text);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat TextFormat</param>
		/// <param name="text">object Text</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_TextWidth(NetOffice.OWC10Api.TextFormat textFormat, object text)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(textFormat, text);
			object returnItem = Invoker.PropertyGet(this, "TextWidth", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_TextWidth
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat TextFormat</param>
		/// <param name="text">object Text</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 TextWidth(NetOffice.OWC10Api.TextFormat textFormat, object text)
		{
			return get_TextWidth(textFormat, text);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="picture">stdole.Picture Picture</param>
		/// <param name="mask">stdole.Picture Mask</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public stdole.Picture get_PictureAlphaBlended(stdole.Picture picture, stdole.Picture mask)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(picture, mask);
			object returnItem = Invoker.PropertyGet(this, "PictureAlphaBlended", paramsArray);
			stdole.Picture newObject = Factory.CreateObjectFromComProxy(this,returnItem) as stdole.Picture;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_PictureAlphaBlended
		/// </summary>
		/// <param name="picture">stdole.Picture Picture</param>
		/// <param name="mask">stdole.Picture Mask</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public stdole.Picture PictureAlphaBlended(stdole.Picture picture, stdole.Picture mask)
		{
			return get_PictureAlphaBlended(picture, mask);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="x">Int32 x</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ScaleX(Int32 x)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(x);
			object returnItem = Invoker.PropertyGet(this, "ScaleX", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_ScaleX
		/// </summary>
		/// <param name="x">Int32 x</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 ScaleX(Int32 x)
		{
			return get_ScaleX(x);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="y">Int32 y</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ScaleY(Int32 y)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(y);
			object returnItem = Invoker.PropertyGet(this, "ScaleY", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_ScaleY
		/// </summary>
		/// <param name="y">Int32 y</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 ScaleY(Int32 y)
		{
			return get_ScaleY(y);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="left">Int32 Left</param>
		/// <param name="top">Int32 Top</param>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		/// <param name="color">Int32 Color</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Rectangle(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, Int32 color)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cx1, cy1, cx2, cy2, left, top, width, height, color);
			Invoker.Method(this, "Rectangle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="x1">Int32 x1</param>
		/// <param name="y1">Int32 y1</param>
		/// <param name="x2">Int32 x2</param>
		/// <param name="y2">Int32 y2</param>
		/// <param name="color">Int32 Color</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Line(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 x1, Int32 y1, Int32 x2, Int32 y2, Int32 color)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cx1, cy1, cx2, cy2, x1, y1, x2, y2, color);
			Invoker.Method(this, "Line", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="left">Int32 Left</param>
		/// <param name="top">Int32 Top</param>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat TextFormat</param>
		/// <param name="text">object Text</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Text(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, NetOffice.OWC10Api.TextFormat textFormat, object text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cx1, cy1, cx2, cy2, left, top, width, height, textFormat, text);
			Invoker.Method(this, "Text", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="left">Int32 Left</param>
		/// <param name="top">Int32 Top</param>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		/// <param name="picture">stdole.Picture Picture</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Picture(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, stdole.Picture picture)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cx1, cy1, cx2, cy2, left, top, width, height, picture);
			Invoker.Method(this, "Picture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="left">Int32 Left</param>
		/// <param name="top">Int32 Top</param>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		/// <param name="picture">stdole.Picture Picture</param>
		/// <param name="mask">stdole.Picture Mask</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void PictureMasked(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, stdole.Picture picture, stdole.Picture mask)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cx1, cy1, cx2, cy2, left, top, width, height, picture, mask);
			Invoker.Method(this, "PictureMasked", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}