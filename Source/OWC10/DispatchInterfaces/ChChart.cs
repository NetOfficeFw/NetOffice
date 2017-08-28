using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ChChart 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ChChart : COMObject
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
                    _type = typeof(ChChart);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ChChart(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ChChart(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChart(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChart(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChart(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChart(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChart() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChChart(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChAxes Axes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChAxes>(this, "Axes", NetOffice.OWC10Api.ChAxes.LateBindingApiWrapperType);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasLegend
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasLegend");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasLegend", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasTitle
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasTitle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasTitle", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 WidthRatio
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WidthRatio");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WidthRatio", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 HeightRatio
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HeightRatio");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HeightRatio", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
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
		public NetOffice.OWC10Api.ChLegend Legend
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChLegend>(this, "Legend", NetOffice.OWC10Api.ChLegend.LateBindingApiWrapperType);
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
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChPlotArea PlotArea
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChPlotArea>(this, "PlotArea", NetOffice.OWC10Api.ChPlotArea.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChTitle Title
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChTitle>(this, "Title", NetOffice.OWC10Api.ChTitle.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double ExtrudeAngle
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "ExtrudeAngle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ExtrudeAngle", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double DirectionalLightRotation
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "DirectionalLightRotation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DirectionalLightRotation", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double DirectionalLightInclination
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "DirectionalLightInclination");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DirectionalLightInclination", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double DirectionalLightIntensity
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "DirectionalLightIntensity");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DirectionalLightIntensity", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double AmbientLightIntensity
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "AmbientLightIntensity");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AmbientLightIntensity", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double LightNormal
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "LightNormal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LightNormal", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartProjectionModeEnum ProjectionMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartProjectionModeEnum>(this, "ProjectionMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ProjectionMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double Inclination
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Inclination");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Inclination", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double Rotation
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Rotation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Rotation", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Perspective
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Perspective");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Perspective", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 GapDepth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GapDepth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GapDepth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasAutoChartDepth
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasAutoChartDepth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasAutoChartDepth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasAutoAspectRatio
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasAutoAspectRatio");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasAutoAspectRatio", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 ChartDepth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ChartDepth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ChartDepth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 AspectRatio
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AspectRatio");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AspectRatio", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double BubbleScale
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "BubbleScale");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BubbleScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		public Int32 HoleSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HoleSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoleSize", value);
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
		public NetOffice.OWC10Api.ChSeriesCollection SeriesCollection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChSeriesCollection>(this, "SeriesCollection", NetOffice.OWC10Api.ChSeriesCollection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartSizeRepresentsEnum SizeRepresents
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartSizeRepresentsEnum>(this, "SizeRepresents");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SizeRepresents", value);
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
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 LayerCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LayerCount");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Top
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Top");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Left
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Left");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Left", value);
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
		[SupportByVersion("OWC10", 1)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
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
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension</param>
		[SupportByVersion("OWC10", 1)]
		public string GetDataReference(NetOffice.OWC10Api.Enums.ChartDimensionsEnum dimension)
		{
			return Factory.ExecuteStringMethodGet(this, "GetDataReference", dimension);
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
		public void DuplicateFormat()
		{
			 Factory.ExecuteMethod(this, "DuplicateFormat");
		}

		#endregion

		#pragma warning restore
	}
}
