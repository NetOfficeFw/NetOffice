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
 	public class _NumberFormat : COMObject
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
                    _type = typeof(_NumberFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _NumberFormat(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		[SupportByVersion("OWC10", 1)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Code
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Code");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Code", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="value">object value</param>
		/// <param name="count">optional Int32 count</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Format(object value, object count)
		{
			return Factory.ExecuteStringPropertyGet(this, "Format", value, count);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Format
		/// </summary>
		/// <param name="value">object value</param>
		/// <param name="count">optional Int32 count</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Format")]
		public string Format(object value, object count)
		{
			return get_Format(value, count);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Format(object value)
		{
			return Factory.ExecuteStringPropertyGet(this, "Format", value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Format
		/// </summary>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Format")]
		public string Format(object value)
		{
			return get_Format(value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_Width(Int32 hDC, object value)
		{
			return Factory.ExecuteInt32PropertyGet(this, "Width", hDC, value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Width
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Width")]
		public Int32 Width(Int32 hDC, object value)
		{
			return get_Width(hDC, value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_Height(Int32 hDC, object value)
		{
			return Factory.ExecuteInt32PropertyGet(this, "Height", hDC, value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Height
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Height")]
		public Int32 Height(Int32 hDC, object value)
		{
			return get_Height(hDC, value);
		}

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
		public void Render(Int32 hDC, Int32 hDCInfo, Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height, Int32 horizontalAlignment, Int32 verticalAlignment, object value)
		{
			 Factory.ExecuteMethod(this, "Render", new object[]{ hDC, hDCInfo, cx1, cy1, cx2, cy2, left, top, width, height, horizontalAlignment, verticalAlignment, value });
		}

		#endregion

		#pragma warning restore
	}
}
