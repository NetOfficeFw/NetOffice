using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Series 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192186.aspx </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Series : COMObject
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
                    _type = typeof(Series);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Series(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Series(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Series(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Series(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Series(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Series(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Series() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Series(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835480.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837155.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839776.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ChartBorder Border
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartBorder>(this, "Border", NetOffice.WordApi.ChartBorder.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840473.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ErrorBars ErrorBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ErrorBars>(this, "ErrorBars", NetOffice.WordApi.ErrorBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839605.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820878.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194755.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197868.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845143.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193005.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837896.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.Interior Interior
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Interior>(this, "Interior", NetOffice.WordApi.Interior.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartFillFormat Fill
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartFillFormat>(this, "Fill", NetOffice.WordApi.ChartFillFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193690.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822949.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192583.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlColorIndex MarkerBackgroundColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlColorIndex>(this, "MarkerBackgroundColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerBackgroundColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836306.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841077.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlColorIndex MarkerForegroundColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlColorIndex>(this, "MarkerForegroundColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerForegroundColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840882.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839091.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlMarkerStyle MarkerStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlMarkerStyle>(this, "MarkerStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196710.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192612.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlChartPictureType PictureType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlChartPictureType>(this, "PictureType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PictureType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822564.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834514.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194258.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835139.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196721.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195596.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195117.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839100.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlBarShape BarShape
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlBarShape>(this, "BarShape");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BarShape", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837212.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822187.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838677.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838082.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838331.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837931.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196557.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.LeaderLines LeaderLines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.LeaderLines>(this, "LeaderLines", NetOffice.WordApi.LeaderLines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839346.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ChartFormat Format
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartFormat>(this, "Format", NetOffice.WordApi.ChartFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192162.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197232.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194786.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838926.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Int32 PlotColorIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PlotColorIndex");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195651.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192178.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlColorIndex InvertColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlColorIndex>(this, "InvertColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "InvertColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227854.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836920.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public object ClearFormats()
		{
			return Factory.ExecuteVariantMethodGet(this, "ClearFormats");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197489.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public object Copy()
		{
			return Factory.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198047.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Word", 14,15,16)]
		public object DataLabels(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "DataLabels", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198047.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object DataLabels()
		{
			return Factory.ExecuteVariantMethodGet(this, "DataLabels");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197208.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public object Delete()
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196738.aspx </remarks>
		/// <param name="direction">NetOffice.WordApi.Enums.XlErrorBarDirection direction</param>
		/// <param name="include">NetOffice.WordApi.Enums.XlErrorBarInclude include</param>
		/// <param name="type">NetOffice.WordApi.Enums.XlErrorBarType type</param>
		/// <param name="amount">optional object amount</param>
		/// <param name="minusValues">optional object minusValues</param>
		[SupportByVersion("Word", 14,15,16)]
		public object ErrorBar(NetOffice.WordApi.Enums.XlErrorBarDirection direction, NetOffice.WordApi.Enums.XlErrorBarInclude include, NetOffice.WordApi.Enums.XlErrorBarType type, object amount, object minusValues)
		{
			return Factory.ExecuteVariantMethodGet(this, "ErrorBar", new object[]{ direction, include, type, amount, minusValues });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196738.aspx </remarks>
		/// <param name="direction">NetOffice.WordApi.Enums.XlErrorBarDirection direction</param>
		/// <param name="include">NetOffice.WordApi.Enums.XlErrorBarInclude include</param>
		/// <param name="type">NetOffice.WordApi.Enums.XlErrorBarType type</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ErrorBar(NetOffice.WordApi.Enums.XlErrorBarDirection direction, NetOffice.WordApi.Enums.XlErrorBarInclude include, NetOffice.WordApi.Enums.XlErrorBarType type)
		{
			return Factory.ExecuteVariantMethodGet(this, "ErrorBar", direction, include, type);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196738.aspx </remarks>
		/// <param name="direction">NetOffice.WordApi.Enums.XlErrorBarDirection direction</param>
		/// <param name="include">NetOffice.WordApi.Enums.XlErrorBarInclude include</param>
		/// <param name="type">NetOffice.WordApi.Enums.XlErrorBarType type</param>
		/// <param name="amount">optional object amount</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ErrorBar(NetOffice.WordApi.Enums.XlErrorBarDirection direction, NetOffice.WordApi.Enums.XlErrorBarInclude include, NetOffice.WordApi.Enums.XlErrorBarType type, object amount)
		{
			return Factory.ExecuteVariantMethodGet(this, "ErrorBar", direction, include, type, amount);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838469.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public object Paste()
		{
			return Factory.ExecuteVariantMethodGet(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193133.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Word", 14,15,16)]
		public object Points(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Points", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193133.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object Points()
		{
			return Factory.ExecuteVariantMethodGet(this, "Points");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836566.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public object Select()
		{
			return Factory.ExecuteVariantMethodGet(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193090.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Word", 14,15,16)]
		public object Trendlines(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Trendlines", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193090.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object Trendlines()
		{
			return Factory.ExecuteVariantMethodGet(this, "Trendlines");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
		{
			 Factory.ExecuteMethod(this, "ApplyCustomType", chartType);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840363.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		/// <param name="showBubbleSize">optional object showBubbleSize</param>
		/// <param name="separator">optional object separator</param>
		[SupportByVersion("Word", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840363.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ApplyDataLabels()
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840363.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ApplyDataLabels(object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840363.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, legendKey);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840363.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, legendKey, autoText);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840363.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, legendKey, autoText, hasLeaderLines);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840363.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840363.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840363.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840363.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840363.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		/// <param name="showBubbleSize">optional object showBubbleSize</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize });
		}

		#endregion

		#pragma warning restore
	}
}
