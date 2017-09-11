using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ChChartSpace 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class ChChartSpace : COMObject
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
                    _type = typeof(ChChartSpace);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ChChartSpace(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ChChartSpace(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartSpace(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartSpace(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartSpace(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartSpace(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartSpace() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChartSpace(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartChartLayoutEnum ChartLayout
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartChartLayoutEnum>(this, "ChartLayout");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ChartLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 ChartWrapCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ChartWrapCount");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ChartWrapCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool EnableEvents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableEvents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableEvents", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasChartSpaceLegend
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasChartSpaceLegend");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasChartSpaceLegend", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 MajorVersion
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MajorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string MinorVersion
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MinorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string BuildNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BuildNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool ScreenUpdating
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ScreenUpdating");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScreenUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChBorder Border
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChBorder>(this, "Border", NetOffice.OWC10Api.ChBorder.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChCharts Charts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChCharts>(this, "Charts", NetOffice.OWC10Api.ChCharts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.MSDATASRCApi.DataSource DataSource
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSDATASRCApi.DataSource>(this, "DataSource", NetOffice.MSDATASRCApi.DataSource.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string DataMember
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataMember");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataMember", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartDataSourceTypeEnum DataSourceType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartDataSourceTypeEnum>(this, "DataSourceType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasChartSpaceTitle
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasChartSpaceTitle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasChartSpaceTitle", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChInterior Interior
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChInterior>(this, "Interior", NetOffice.OWC10Api.ChInterior.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChLegend ChartSpaceLegend
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChLegend>(this, "ChartSpaceLegend", NetOffice.OWC10Api.ChLegend.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public object Selection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Selection");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartSelectionsEnum SelectionType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartSelectionsEnum>(this, "SelectionType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartSelectionMarksEnum HasSelectionMarks
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartSelectionMarksEnum>(this, "HasSelectionMarks");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HasSelectionMarks", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayPropertyToolbox
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayPropertyToolbox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayPropertyToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChTitle ChartSpaceTitle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChTitle>(this, "ChartSpaceTitle", NetOffice.OWC10Api.ChTitle.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string XMLData
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XMLData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XMLData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Constants
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Constants");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool CanUndo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CanUndo");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowLayoutEvents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowLayoutEvents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowLayoutEvents", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowRenderEvents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowRenderEvents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowRenderEvents", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowPointRenderEvents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowPointRenderEvents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowPointRenderEvents", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Enabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string RevisionNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RevisionNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double PrintQuality3D
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "PrintQuality3D");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintQuality3D", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayScreenTips
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayScreenTips");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayScreenTips", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string ConnectionString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConnectionString");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConnectionString", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string CommandText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandText", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public object InternalPivotTable
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "InternalPivotTable");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasSeriesByRows
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasSeriesByRows");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasSeriesByRows", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartPlotAggregatesEnum PlotAllAggregates
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartPlotAggregatesEnum>(this, "PlotAllAggregates");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PlotAllAggregates", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasMultipleCharts
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasMultipleCharts");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasMultipleCharts", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayFieldList
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayFieldList");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayFieldList", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasPassiveAlerts
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasPassiveAlerts");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasPassiveAlerts", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string DataSourceName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataSourceName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataSourceName", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayFieldButtons
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayFieldButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayFieldButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object SelectionList
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SelectionList");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasPlotDetails
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasPlotDetails");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasPlotDetails", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowScreenTipEvents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowScreenTipEvents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowScreenTipEvents", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.OCCommands Commands
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.OCCommands>(this, "Commands", NetOffice.OWC10Api.OCCommands.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowPropertyToolbox
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowPropertyToolbox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowPropertyToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowGrouping
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowGrouping");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowGrouping", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowFiltering
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFiltering");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowFiltering", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Top
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Left
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Bottom
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Bottom");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Right
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Right");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasUnifiedScales
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasUnifiedScales");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasUnifiedScales", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayToolbar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayToolbar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayToolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.MSComctlLibApi.IToolbar Toolbar
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IToolbar>(this, "Toolbar");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ViewOnlyMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ViewOnlyMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool IsDirty
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsDirty");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsDirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_International(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "International", index);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_International
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1), Redirect("get_International")]
		public object International(object index)
		{
			return get_International(index);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.OWCLanguageSettings LanguageSettings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.OWCLanguageSettings>(this, "LanguageSettings", NetOffice.OWC10Api.OWCLanguageSettings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasRuntimeSelection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasRuntimeSelection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasRuntimeSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayBranding
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayBranding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayBranding", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayOfficeLogo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayOfficeLogo");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayOfficeLogo", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartSelectionsEnum ObjectType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartSelectionsEnum>(this, "ObjectType");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void BuildLitChart()
		{
			 Factory.ExecuteMethod(this, "BuildLitChart");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">string filename</param>
		[SupportByVersion("OWC10", 1)]
		public void Load(string filename)
		{
			 Factory.ExecuteMethod(this, "Load", filename);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="iTopic">Int32 iTopic</param>
		[SupportByVersion("OWC10", 1)]
		public void ShowHelp(Int32 iTopic)
		{
			 Factory.ExecuteMethod(this, "ShowHelp", iTopic);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = chart.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = -1</param>
		/// <param name="height">optional Int32 Height = -1</param>
		[SupportByVersion("OWC10", 1)]
		public void ExportPicture(object filename, object filterName, object width, object height)
		{
			 Factory.ExecuteMethod(this, "ExportPicture", filename, filterName, width, height);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ExportPicture()
		{
			 Factory.ExecuteMethod(this, "ExportPicture");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = chart.gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ExportPicture(object filename)
		{
			 Factory.ExecuteMethod(this, "ExportPicture", filename);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = chart.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ExportPicture(object filename, object filterName)
		{
			 Factory.ExecuteMethod(this, "ExportPicture", filename, filterName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = chart.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = -1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ExportPicture(object filename, object filterName, object width)
		{
			 Factory.ExecuteMethod(this, "ExportPicture", filename, filterName, width);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void BeginUndo()
		{
			 Factory.ExecuteMethod(this, "BeginUndo");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void EndUndo()
		{
			 Factory.ExecuteMethod(this, "EndUndo");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Undo()
		{
			 Factory.ExecuteMethod(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("OWC10", 1)]
		public object RangeFromPoint(Int32 x, Int32 y)
		{
			return Factory.ExecuteVariantMethodGet(this, "RangeFromPoint", x, y);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		/// <param name="dataSourceIndex">Int32 dataSourceIndex</param>
		/// <param name="dataReference">optional object dataReference</param>
		[SupportByVersion("OWC10", 1)]
		public void SetData(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension, Int32 dataSourceIndex, object dataReference)
		{
			 Factory.ExecuteMethod(this, "SetData", dimension, dataSourceIndex, dataReference);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		/// <param name="dataSourceIndex">Int32 dataSourceIndex</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetData(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension, Int32 dataSourceIndex)
		{
			 Factory.ExecuteMethod(this, "SetData", dimension, dataSourceIndex);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dz">NetOffice.OWC10Api.Enums.ChartDropZonesEnum dz</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChDropZone DropZones(NetOffice.OWC10Api.Enums.ChartDropZonesEnum dz)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.ChDropZone>(this, "DropZones", NetOffice.OWC10Api.ChDropZone.LateBindingApiWrapperType, dz);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="punk">object punk</param>
		/// <param name="lPos">Int32 lPos</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void FieldListAddTo(object punk, Int32 lPos)
		{
			 Factory.ExecuteMethod(this, "FieldListAddTo", punk, lPos);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void LocateDataSource()
		{
			 Factory.ExecuteMethod(this, "LocateDataSource");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="menu">object menu</param>
		[SupportByVersion("OWC10", 1)]
		public void ShowContextMenu(Int32 x, Int32 y, object menu)
		{
			 Factory.ExecuteMethod(this, "ShowContextMenu", x, y, menu);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = -1</param>
		/// <param name="height">optional Int32 Height = -1</param>
		[SupportByVersion("OWC10", 1)]
		public object GetPicture(object filterName, object width, object height)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetPicture", filterName, width, height);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public object GetPicture()
		{
			return Factory.ExecuteVariantMethodGet(this, "GetPicture");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filterName">optional string FilterName = gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public object GetPicture(object filterName)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetPicture", filterName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = -1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public object GetPicture(object filterName, object width)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetPicture", filterName, width);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dataReference">string dataReference</param>
		/// <param name="seriesByRows">optional bool SeriesByRows = false</param>
		[SupportByVersion("OWC10", 1)]
		public void SetSpreadsheetData(string dataReference, object seriesByRows)
		{
			 Factory.ExecuteMethod(this, "SetSpreadsheetData", dataReference, seriesByRows);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dataReference">string dataReference</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetSpreadsheetData(string dataReference)
		{
			 Factory.ExecuteMethod(this, "SetSpreadsheetData", dataReference);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Repaint()
		{
			 Factory.ExecuteMethod(this, "Repaint");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 GetDataSourceIndex(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetDataSourceIndex", dimension);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void ClearUndo()
		{
			 Factory.ExecuteMethod(this, "ClearUndo");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void OkToBindToControlByName()
		{
			 Factory.ExecuteMethod(this, "OkToBindToControlByName");
		}

		#endregion

		#pragma warning restore
	}
}
