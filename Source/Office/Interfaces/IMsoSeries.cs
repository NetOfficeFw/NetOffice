using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// Interface IMsoSeries 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IMsoSeries : COMObject
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
                    _type = typeof(IMsoSeries);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IMsoSeries(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMsoSeries(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoSeries(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoSeries(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoSeries(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoSeries(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoSeries() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoSeries(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlAxisGroup AxisGroup
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlAxisGroup>(this, "AxisGroup");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AxisGroup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoBorder Border
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoBorder>(this, "Border", NetOffice.OfficeApi.IMsoBorder.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoErrorBars ErrorBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoErrorBars>(this, "ErrorBars", NetOffice.OfficeApi.IMsoErrorBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 Explosion
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Explosion");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Explosion", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public string Formula
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Formula");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Formula", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public string FormulaLocal
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FormulaLocal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormulaLocal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public string FormulaR1C1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FormulaR1C1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormulaR1C1", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public string FormulaR1C1Local
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FormulaR1C1Local");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormulaR1C1Local", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool HasDataLabels
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasDataLabels");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasDataLabels", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool HasErrorBars
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasErrorBars");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasErrorBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoInterior Interior
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoInterior>(this, "Interior", NetOffice.OfficeApi.IMsoInterior.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.ChartFillFormat Fill
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ChartFillFormat>(this, "Fill", NetOffice.OfficeApi.ChartFillFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlColorIndex MarkerBackgroundColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlColorIndex>(this, "MarkerBackgroundColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerBackgroundColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlColorIndex MarkerForegroundColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlColorIndex>(this, "MarkerForegroundColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerForegroundColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlMarkerStyle MarkerStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlMarkerStyle>(this, "MarkerStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlChartPictureType PictureType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlChartPictureType>(this, "PictureType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PictureType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public Double PictureUnit
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "PictureUnit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 PlotOrder
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PlotOrder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PlotOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlChartType ChartType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlChartType>(this, "ChartType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ChartType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public object Values
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Values");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Values", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public object XValues
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "XValues");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "XValues", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public object BubbleSizes
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BubbleSizes");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BubbleSizes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlBarShape BarShape
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlBarShape>(this, "BarShape");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BarShape", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool ApplyPictToSides
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyPictToSides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyPictToSides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool ApplyPictToFront
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyPictToFront");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyPictToFront", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool ApplyPictToEnd
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyPictToEnd");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyPictToEnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool Has3DEffect
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Has3DEffect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Has3DEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool HasLeaderLines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasLeaderLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasLeaderLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoLeaderLines LeaderLines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoLeaderLines>(this, "LeaderLines", NetOffice.OfficeApi.IMsoLeaderLines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoChartFormat Format
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartFormat>(this, "Format", NetOffice.OfficeApi.IMsoChartFormat.LateBindingApiWrapperType);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 PlotColorIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PlotColorIndex");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 InvertColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "InvertColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InvertColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.Enums.XlColorIndex InvertColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlColorIndex>(this, "InvertColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "InvertColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 15, 16)]
		public bool IsFiltered
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsFiltered");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsFiltered", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 12,14,15,16)]
		public object _ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines)
		{
			return Factory.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, iMsoLegendKey, autoText, hasLeaderLines);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object _ApplyDataLabels()
		{
			return Factory.ExecuteVariantMethodGet(this, "_ApplyDataLabels");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object _ApplyDataLabels(object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object _ApplyDataLabels(object type, object iMsoLegendKey)
		{
			return Factory.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, iMsoLegendKey);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		/// <param name="autoText">optional object autoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object _ApplyDataLabels(object type, object iMsoLegendKey, object autoText)
		{
			return Factory.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, iMsoLegendKey, autoText);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public object ClearFormats()
		{
			return Factory.ExecuteVariantMethodGet(this, "ClearFormats");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public object Copy()
		{
			return Factory.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public object DataLabels(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "DataLabels", index);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object DataLabels()
		{
			return Factory.ExecuteVariantMethodGet(this, "DataLabels");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public object Delete()
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.OfficeApi.Enums.XlErrorBarDirection direction</param>
		/// <param name="include">NetOffice.OfficeApi.Enums.XlErrorBarInclude include</param>
		/// <param name="type">NetOffice.OfficeApi.Enums.XlErrorBarType type</param>
		/// <param name="amount">optional object amount</param>
		/// <param name="minusValues">optional object minusValues</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public object ErrorBar(NetOffice.OfficeApi.Enums.XlErrorBarDirection direction, NetOffice.OfficeApi.Enums.XlErrorBarInclude include, NetOffice.OfficeApi.Enums.XlErrorBarType type, object amount, object minusValues)
		{
			return Factory.ExecuteVariantMethodGet(this, "ErrorBar", new object[]{ direction, include, type, amount, minusValues });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.OfficeApi.Enums.XlErrorBarDirection direction</param>
		/// <param name="include">NetOffice.OfficeApi.Enums.XlErrorBarInclude include</param>
		/// <param name="type">NetOffice.OfficeApi.Enums.XlErrorBarType type</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ErrorBar(NetOffice.OfficeApi.Enums.XlErrorBarDirection direction, NetOffice.OfficeApi.Enums.XlErrorBarInclude include, NetOffice.OfficeApi.Enums.XlErrorBarType type)
		{
			return Factory.ExecuteVariantMethodGet(this, "ErrorBar", direction, include, type);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.OfficeApi.Enums.XlErrorBarDirection direction</param>
		/// <param name="include">NetOffice.OfficeApi.Enums.XlErrorBarInclude include</param>
		/// <param name="type">NetOffice.OfficeApi.Enums.XlErrorBarType type</param>
		/// <param name="amount">optional object amount</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ErrorBar(NetOffice.OfficeApi.Enums.XlErrorBarDirection direction, NetOffice.OfficeApi.Enums.XlErrorBarInclude include, NetOffice.OfficeApi.Enums.XlErrorBarType type, object amount)
		{
			return Factory.ExecuteVariantMethodGet(this, "ErrorBar", direction, include, type, amount);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public object Paste()
		{
			return Factory.ExecuteVariantMethodGet(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public object Points(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Points", index);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object Points()
		{
			return Factory.ExecuteVariantMethodGet(this, "Points");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public object Select()
		{
			return Factory.ExecuteVariantMethodGet(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public object Trendlines(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Trendlines", index);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object Trendlines()
		{
			return Factory.ExecuteVariantMethodGet(this, "Trendlines");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
		{
			return Factory.ExecuteInt32MethodGet(this, "ApplyCustomType", chartType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		/// <param name="showBubbleSize">optional object showBubbleSize</param>
		/// <param name="separator">optional object separator</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ApplyDataLabels()
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ApplyDataLabels(object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ApplyDataLabels(object type, object iMsoLegendKey)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, iMsoLegendKey);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ApplyDataLabels(object type, object iMsoLegendKey, object autoText)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, iMsoLegendKey, autoText);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, iMsoLegendKey, autoText, hasLeaderLines);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		/// <param name="showBubbleSize">optional object showBubbleSize</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize });
		}

		#endregion

		#pragma warning restore
	}
}
