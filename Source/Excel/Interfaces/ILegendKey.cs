using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface ILegendKey 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class ILegendKey : COMObject
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
                    _type = typeof(ILegendKey);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ILegendKey(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ILegendKey(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILegendKey(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILegendKey(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILegendKey(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILegendKey(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILegendKey() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILegendKey(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Border Border
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Border>(this, "Border", NetOffice.ExcelApi.Border.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Interior Interior
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Interior>(this, "Interior", NetOffice.ExcelApi.Interior.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ChartFillFormat Fill
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ChartFillFormat>(this, "Fill", NetOffice.ExcelApi.ChartFillFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool InvertIfNegative
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InvertIfNegative");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InvertIfNegative", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 MarkerBackgroundColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MarkerBackgroundColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarkerBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlColorIndex MarkerBackgroundColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlColorIndex>(this, "MarkerBackgroundColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerBackgroundColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 MarkerForegroundColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MarkerForegroundColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarkerForegroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlColorIndex MarkerForegroundColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlColorIndex>(this, "MarkerForegroundColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerForegroundColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 MarkerSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MarkerSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarkerSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlMarkerStyle MarkerStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlMarkerStyle>(this, "MarkerStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PictureType
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PictureType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PictureUnit
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PictureUnit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Smooth
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Smooth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Smooth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double Left
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double Top
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double Width
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Width");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double Height
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Height");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Shadow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Shadow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Shadow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Double PictureUnit2
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "PictureUnit2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureUnit2", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ChartFormat Format
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ChartFormat>(this, "Format", NetOffice.ExcelApi.ChartFormat.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ClearFormats()
		{
			return Factory.ExecuteVariantMethodGet(this, "ClearFormats");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Delete()
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Select()
		{
			return Factory.ExecuteVariantMethodGet(this, "Select");
		}

		#endregion

		#pragma warning restore
	}
}
