using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ChSeries 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ChSeries : COMObject
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
                    _type = typeof(ChSeries);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ChSeries(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ChSeries(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChSeries(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChSeries(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChSeries(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChSeries(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChSeries() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChSeries(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		public NetOffice.OWC10Api.ChDataLabelsCollection DataLabelsCollection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChDataLabelsCollection>(this, "DataLabelsCollection", NetOffice.OWC10Api.ChDataLabelsCollection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChPoints Points
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChPoints>(this, "Points", NetOffice.OWC10Api.ChPoints.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Thickness
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Thickness");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Thickness", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChErrorBarsCollection ErrorBarsCollection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChErrorBarsCollection>(this, "ErrorBarsCollection", NetOffice.OWC10Api.ChErrorBarsCollection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Index", value);
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
		public NetOffice.OWC10Api.ChLine Line
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChLine>(this, "Line", NetOffice.OWC10Api.ChLine.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChMarker Marker
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChMarker>(this, "Marker", NetOffice.OWC10Api.ChMarker.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChChart Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChChart>(this, "Parent", NetOffice.OWC10Api.ChChart.LateBindingApiWrapperType);
			}
		}

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
		/// Get
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.ChScaling get_Scalings(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChScaling>(this, "Scalings", NetOffice.OWC10Api.ChScaling.LateBindingApiWrapperType, dimension);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Scalings
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Scalings")]
		public NetOffice.OWC10Api.ChScaling Scalings(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension)
		{
			return get_Scalings(dimension);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChTrendlines Trendlines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChTrendlines>(this, "Trendlines", NetOffice.OWC10Api.ChTrendlines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartChartTypeEnum Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartChartTypeEnum>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 ZOrder
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ZOrder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ZOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public object PivotObject
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "PivotObject");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChFormatMap FormatMap
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChFormatMap>(this, "FormatMap", NetOffice.OWC10Api.ChFormatMap.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string TipText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TipText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TipText", value);
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
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 LayerIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LayerIndex");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 TypeFlags
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TypeFlags");
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
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public string GetDataReference(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension)
		{
			return Factory.ExecuteStringMethodGet(this, "GetDataReference", dimension);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public Int32 GetDataSourceIndex(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetDataSourceIndex", dimension);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		/// <param name="dataSourceIndex">object dataSourceIndex</param>
		/// <param name="dataReference">object dataReference</param>
		[SupportByVersion("OWC10", 1)]
		public void GetData(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension, out object dataSourceIndex, out object dataReference)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			dataSourceIndex = null;
			dataReference = null;
			object[] paramsArray = Invoker.ValidateParamsArray(dimension, dataSourceIndex, dataReference);
			Invoker.Method(this, "GetData", paramsArray, modifiers);
			dataSourceIndex = (object)paramsArray[1];
			dataReference = (object)paramsArray[2];
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
		/// <param name="useNewScaling">optional bool UseNewScaling = false</param>
		[SupportByVersion("OWC10", 1)]
		public void Ungroup(object useNewScaling)
		{
			 Factory.ExecuteMethod(this, "Ungroup", useNewScaling);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Ungroup()
		{
			 Factory.ExecuteMethod(this, "Ungroup");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="series">NetOffice.OWC10Api.ChSeries series</param>
		[SupportByVersion("OWC10", 1)]
		public void Group(NetOffice.OWC10Api.ChSeries series)
		{
			 Factory.ExecuteMethod(this, "Group", series);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="xvalue">object xvalue</param>
		/// <param name="yvalue">object yvalue</param>
		/// <param name="zvalue">optional object zvalue</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Coordinate ValueToPoint(object xvalue, object yvalue, object zvalue)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.Coordinate>(this, "ValueToPoint", NetOffice.OWC10Api.Coordinate.LateBindingApiWrapperType, xvalue, yvalue, zvalue);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="xvalue">object xvalue</param>
		/// <param name="yvalue">object yvalue</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Coordinate ValueToPoint(object xvalue, object yvalue)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.Coordinate>(this, "ValueToPoint", NetOffice.OWC10Api.Coordinate.LateBindingApiWrapperType, xvalue, yvalue);
		}

		#endregion

		#pragma warning restore
	}
}
