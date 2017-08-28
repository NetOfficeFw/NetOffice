using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// Interface ChartPoint 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class ChartPoint : COMObject
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
                    _type = typeof(ChartPoint);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ChartPoint(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ChartPoint(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartPoint(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartPoint(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartPoint(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartPoint(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartPoint() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChartPoint(string progId) : base(progId)
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
		public NetOffice.OfficeApi.IMsoDataLabel DataLabel
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoDataLabel>(this, "DataLabel", NetOffice.OfficeApi.IMsoDataLabel.LateBindingApiWrapperType);
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
		public bool HasDataLabel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasDataLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasDataLabel", value);
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
		public bool SecondaryPlot
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SecondaryPlot");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SecondaryPlot", value);
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
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 14,15,16)]
		public Double Height
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Height");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 14,15,16)]
		public Double Width
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Width");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 14,15,16)]
		public Double Left
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 14,15,16)]
		public Double Top
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Top");
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
		[SupportByVersion("Office", 12,14,15,16)]
		public object Delete()
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete");
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
		[SupportByVersion("Office", 12,14,15,16)]
		public object Select()
		{
			return Factory.ExecuteVariantMethodGet(this, "Select");
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

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <param name="loc">NetOffice.OfficeApi.Enums.XlPieSliceLocation loc</param>
		/// <param name="index">optional NetOffice.OfficeApi.Enums.XlPieSliceIndex Index = 2</param>
		[SupportByVersion("Office", 14,15,16)]
		public Double PieSliceLocation(NetOffice.OfficeApi.Enums.XlPieSliceLocation loc, object index)
		{
			return Factory.ExecuteDoubleMethodGet(this, "PieSliceLocation", loc, index);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <param name="loc">NetOffice.OfficeApi.Enums.XlPieSliceLocation loc</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public Double PieSliceLocation(NetOffice.OfficeApi.Enums.XlPieSliceLocation loc)
		{
			return Factory.ExecuteDoubleMethodGet(this, "PieSliceLocation", loc);
		}

		#endregion

		#pragma warning restore
	}
}
