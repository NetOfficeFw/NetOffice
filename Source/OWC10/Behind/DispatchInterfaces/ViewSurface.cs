using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface ViewSurface 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ViewSurface : COMObject, NetOffice.OWC10Api.ViewSurface
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
                    _contractType = typeof(NetOffice.OWC10Api.ViewSurface);
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
                    _type = typeof(ViewSurface);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ViewSurface() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 hDC
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "hDC");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 hDCInfo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "hDCInfo");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="color">Int32 color</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_AlphaBlend(Int32 color)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AlphaBlend", color);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_AlphaBlend
		/// </summary>
		/// <param name="color">Int32 color</param>
		[SupportByVersion("OWC10", 1), Redirect("get_AlphaBlend")]
		public virtual Int32 AlphaBlend(Int32 color)
		{
			return get_AlphaBlend(color);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat textFormat</param>
		/// <param name="text">object text</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_TextHeight(NetOffice.OWC10Api.TextFormat textFormat, object text)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TextHeight", textFormat, text);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_TextHeight
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat textFormat</param>
		/// <param name="text">object text</param>
		[SupportByVersion("OWC10", 1), Redirect("get_TextHeight")]
		public virtual Int32 TextHeight(NetOffice.OWC10Api.TextFormat textFormat, object text)
		{
			return get_TextHeight(textFormat, text);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat textFormat</param>
		/// <param name="text">object text</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_TextWidth(NetOffice.OWC10Api.TextFormat textFormat, object text)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TextWidth", textFormat, text);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_TextWidth
		/// </summary>
		/// <param name="textFormat">NetOffice.OWC10Api.TextFormat textFormat</param>
		/// <param name="text">object text</param>
		[SupportByVersion("OWC10", 1), Redirect("get_TextWidth")]
		public virtual Int32 TextWidth(NetOffice.OWC10Api.TextFormat textFormat, object text)
		{
			return get_TextWidth(textFormat, text);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="picture">stdole.Picture picture</param>
		/// <param name="mask">stdole.Picture mask</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false), NativeResult]
		public virtual stdole.Picture get_PictureAlphaBlended(stdole.Picture picture, stdole.Picture mask)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(picture, mask);
			object returnItem = Invoker.PropertyGet(this, "PictureAlphaBlended", paramsArray);
            return returnItem as stdole.Picture;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_PictureAlphaBlended
		/// </summary>
		/// <param name="picture">stdole.Picture picture</param>
		/// <param name="mask">stdole.Picture mask</param>
		[SupportByVersion("OWC10", 1), Redirect("get_PictureAlphaBlended")]
		public virtual stdole.Picture PictureAlphaBlended(stdole.Picture picture, stdole.Picture mask)
		{
			return get_PictureAlphaBlended(picture, mask);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="x">Int32 x</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_ScaleX(Int32 x)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ScaleX", x);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_ScaleX
		/// </summary>
		/// <param name="x">Int32 x</param>
		[SupportByVersion("OWC10", 1), Redirect("get_ScaleX")]
		public virtual Int32 ScaleX(Int32 x)
		{
			return get_ScaleX(x);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_ScaleY(Int32 y)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ScaleY", y);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_ScaleY
		/// </summary>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("OWC10", 1), Redirect("get_ScaleY")]
		public virtual Int32 ScaleY(Int32 y)
		{
			return get_ScaleY(y);
		}

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
		public virtual void Rectangle(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, Int32 color)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rectangle", new object[]{ cx1, cy1, cx2, cy2, left, top, width, height, color });
		}

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
		public virtual void Line(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 x1, Int32 y1, Int32 x2, Int32 y2, Int32 color)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Line", new object[]{ cx1, cy1, cx2, cy2, x1, y1, x2, y2, color });
		}

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
		public virtual void Text(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, NetOffice.OWC10Api.TextFormat textFormat, object text)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Text", new object[]{ cx1, cy1, cx2, cy2, left, top, width, height, textFormat, text });
		}

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
		public virtual void Picture(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, stdole.Picture picture)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Picture", new object[]{ cx1, cy1, cx2, cy2, left, top, width, height, picture });
		}

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
		public virtual void PictureMasked(Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, stdole.Picture picture, stdole.Picture mask)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PictureMasked", new object[]{ cx1, cy1, cx2, cy2, left, top, width, height, picture, mask });
		}

		#endregion

		#pragma warning restore
	}
}

