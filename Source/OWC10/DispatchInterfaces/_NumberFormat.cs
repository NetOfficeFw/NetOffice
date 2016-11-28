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
	/// DispatchInterface _NumberFormat 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _NumberFormat : COMObject
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
                    _type = typeof(_NumberFormat);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _NumberFormat(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NumberFormat(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NumberFormat(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NumberFormat(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NumberFormat(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NumberFormat() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NumberFormat(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Code
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Code", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Code", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="value">object Value</param>
		/// <param name="count">optional Int32 Count</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Format(object value, object count)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(value, count);
			object returnItem = Invoker.PropertyGet(this, "Format", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Format
		/// </summary>
		/// <param name="value">object Value</param>
		/// <param name="count">optional Int32 Count</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Format(object value, object count)
		{
			return get_Format(value, count);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="value">object Value</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Format(object value)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(value);
			object returnItem = Invoker.PropertyGet(this, "Format", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Format
		/// </summary>
		/// <param name="value">object Value</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Format(object value)
		{
			return get_Format(value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object Value</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_Width(Int32 hDC, object value)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(hDC, value);
			object returnItem = Invoker.PropertyGet(this, "Width", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Width
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object Value</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 Width(Int32 hDC, object value)
		{
			return get_Width(hDC, value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object Value</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_Height(Int32 hDC, object value)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(hDC, value);
			object returnItem = Invoker.PropertyGet(this, "Height", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Height
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object Value</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 Height(Int32 hDC, object value)
		{
			return get_Height(hDC, value);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="hDCInfo">Int32 hDCInfo</param>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="left">Int32 Left</param>
		/// <param name="top">Int32 Top</param>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		/// <param name="horizontalAlignment">Int32 HorizontalAlignment</param>
		/// <param name="verticalAlignment">Int32 VerticalAlignment</param>
		/// <param name="value">object Value</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Render(Int32 hDC, Int32 hDCInfo, Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, Int32 horizontalAlignment, Int32 verticalAlignment, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hDC, hDCInfo, cx1, cy1, cx2, cy2, left, top, width, height, horizontalAlignment, verticalAlignment, value);
			Invoker.Method(this, "Render", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}