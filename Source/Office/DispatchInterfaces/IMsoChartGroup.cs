using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface IMsoChartGroup 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IMsoChartGroup : COMObject
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
                    _type = typeof(IMsoChartGroup);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IMsoChartGroup(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMsoChartGroup(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChartGroup(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChartGroup(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChartGroup(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChartGroup(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChartGroup() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChartGroup(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 AxisGroup
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AxisGroup");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AxisGroup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 DoughnutHoleSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DoughnutHoleSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DoughnutHoleSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoDownBars DownBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoDownBars>(this, "DownBars", NetOffice.OfficeApi.IMsoDownBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoDropLines DropLines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoDropLines>(this, "DropLines", NetOffice.OfficeApi.IMsoDropLines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 FirstSliceAngle
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FirstSliceAngle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FirstSliceAngle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 GapWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GapWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GapWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool HasDropLines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasDropLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasDropLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool HasHiLoLines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasHiLoLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasHiLoLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool HasRadarAxisLabels
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasRadarAxisLabels");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasRadarAxisLabels", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool HasSeriesLines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasSeriesLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasSeriesLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool HasUpDownBars
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasUpDownBars");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasUpDownBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoHiLoLines HiLoLines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoHiLoLines>(this, "HiLoLines", NetOffice.OfficeApi.IMsoHiLoLines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 Overlap
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Overlap");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Overlap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16), ProxyResult]
		public object RadarAxisLabels
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "RadarAxisLabels");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoSeriesLines SeriesLines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoSeriesLines>(this, "SeriesLines", NetOffice.OfficeApi.IMsoSeriesLines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 SubType
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SubType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SubType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 Type
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Type");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoUpBars UpBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoUpBars>(this, "UpBars", NetOffice.OfficeApi.IMsoUpBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool VaryByCategories
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "VaryByCategories");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "VaryByCategories", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlSizeRepresents SizeRepresents
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlSizeRepresents>(this, "SizeRepresents");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SizeRepresents", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 BubbleScale
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BubbleScale");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BubbleScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool ShowNegativeBubbles
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowNegativeBubbles");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowNegativeBubbles", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlChartSplitType SplitType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlChartSplitType>(this, "SplitType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SplitType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public object SplitValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SplitValue");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "SplitValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 SecondPlotSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SecondPlotSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SecondPlotSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool Has3DShading
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Has3DShading");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Has3DShading", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Office", 14,15,16), ProxyResult]
		public object Application
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Office", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public object SeriesCollection(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "SeriesCollection", index);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object SeriesCollection()
		{
			return Factory.ExecuteVariantMethodGet(this, "SeriesCollection");
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Office", 15, 16)]
		public object CategoryCollection(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "CategoryCollection", index);
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 15, 16)]
		public object CategoryCollection()
		{
			return Factory.ExecuteVariantMethodGet(this, "CategoryCollection");
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Office", 15, 16)]
		public object FullCategoryCollection(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "FullCategoryCollection", index);
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 15, 16)]
		public object FullCategoryCollection()
		{
			return Factory.ExecuteVariantMethodGet(this, "FullCategoryCollection");
		}

		#endregion

		#pragma warning restore
	}
}
