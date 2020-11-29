using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface ChartGroup 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup"/> </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ChartGroup : COMObject
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
                    _type = typeof(ChartGroup);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ChartGroup(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ChartGroup(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartGroup(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartGroup(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartGroup(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartGroup(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartGroup() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartGroup(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.AxisGroup"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlAxisGroup AxisGroup
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlAxisGroup>(this, "AxisGroup");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AxisGroup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.DoughnutHoleSize"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.DownBars"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.DownBars DownBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.DownBars>(this, "DownBars", NetOffice.WordApi.DownBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.DropLines"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.DropLines DropLines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.DropLines>(this, "DropLines", NetOffice.WordApi.DropLines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.FirstSliceAngle"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.GapWidth"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.HasDropLines"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.HasHiLoLines"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.HasRadarAxisLabels"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.HasSeriesLines"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.HasUpDownBars"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.HiLoLines"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.HiLoLines HiLoLines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HiLoLines>(this, "HiLoLines", NetOffice.WordApi.HiLoLines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.Index"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.Overlap"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.RadarAxisLabels"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.TickLabels RadarAxisLabels
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TickLabels>(this, "RadarAxisLabels", NetOffice.WordApi.TickLabels.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.SeriesLines"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.SeriesLines SeriesLines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SeriesLines>(this, "SeriesLines", NetOffice.WordApi.SeriesLines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.UpBars"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.UpBars UpBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.UpBars>(this, "UpBars", NetOffice.WordApi.UpBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.VaryByCategories"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.SizeRepresents"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlSizeRepresents SizeRepresents
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlSizeRepresents>(this, "SizeRepresents");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SizeRepresents", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.BubbleScale"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.ShowNegativeBubbles"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.SplitType"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlChartSplitType SplitType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlChartSplitType>(this, "SplitType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SplitType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.SplitValue"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.SecondPlotSize"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.Has3DShading"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.Application"/> </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		public object Application
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.Creator"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.Parent"/> </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
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
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.SeriesCollection"/> </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Word", 14,15,16)]
		public object SeriesCollection(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "SeriesCollection", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ChartGroup.SeriesCollection"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object SeriesCollection()
		{
			return Factory.ExecuteVariantMethodGet(this, "SeriesCollection");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.chartgroup.categorycollection"/> </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Word", 15, 16)]
		public object CategoryCollection(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "CategoryCollection", index);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.chartgroup.categorycollection"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public object CategoryCollection()
		{
			return Factory.ExecuteVariantMethodGet(this, "CategoryCollection");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.chartgroup.fullcategorycollection"/> </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Word", 15, 16)]
		public object FullCategoryCollection(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "FullCategoryCollection", index);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.chartgroup.fullcategorycollection"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public object FullCategoryCollection()
		{
			return Factory.ExecuteVariantMethodGet(this, "FullCategoryCollection");
		}

		#endregion

		#pragma warning restore
	}
}
