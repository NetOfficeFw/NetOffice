using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ChAxis 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ChAxis : COMObject
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
                    _type = typeof(ChAxis);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ChAxis(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ChAxis(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChAxis(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChAxis(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChAxis(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChAxis(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChAxis() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChAxis(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartAxisCrossesEnum Crosses
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartAxisCrossesEnum>(this, "Crosses");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Crosses", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double CrossesAtValue
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "CrossesAtValue");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CrossesAtValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChAxis CrossingAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChAxis>(this, "CrossingAxis", NetOffice.OWC10Api.ChAxis.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "CrossingAxis", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChFont Font
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChFont>(this, "Font", NetOffice.OWC10Api.ChFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasAutoMajorUnit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasAutoMajorUnit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasAutoMajorUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasAutoMinorUnit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasAutoMinorUnit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasAutoMinorUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasMajorGridlines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasMajorGridlines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasMajorGridlines", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasMinorGridlines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasMinorGridlines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasMinorGridlines", value);
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
		public NetOffice.OWC10Api.ChGridlines MajorGridlines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChGridlines>(this, "MajorGridlines", NetOffice.OWC10Api.ChGridlines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartTickMarkEnum MajorTickMarks
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartTickMarkEnum>(this, "MajorTickMarks");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MajorTickMarks", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double MajorUnit
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "MajorUnit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MajorUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChGridlines MinorGridlines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChGridlines>(this, "MinorGridlines", NetOffice.OWC10Api.ChGridlines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartTickMarkEnum MinorTickMarks
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartTickMarkEnum>(this, "MinorTickMarks");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MinorTickMarks", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double MinorUnit
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "MinorUnit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinorUnit", value);
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
		public NetOffice.OWC10Api.Enums.ChartAxisPositionEnum Position
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartAxisPositionEnum>(this, "Position");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Position", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChScaling Scaling
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChScaling>(this, "Scaling", NetOffice.OWC10Api.ChScaling.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasTickLabels
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasTickLabels");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasTickLabels", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 TickLabelSpacing
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TickLabelSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TickLabelSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 TickMarkSpacing
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TickMarkSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TickMarkSpacing", value);
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
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartAxisTypeEnum Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartAxisTypeEnum>(this, "Type");
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string NumberFormat
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NumberFormat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NumberFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartAxisGroupingEnum GroupingType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartAxisGroupingEnum>(this, "GroupingType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "GroupingType", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartAxisUnitTypeEnum TickLabelUnitType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartAxisUnitTypeEnum>(this, "TickLabelUnitType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TickLabelUnitType", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartAxisUnitTypeEnum TickMarkUnitType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartAxisUnitTypeEnum>(this, "TickMarkUnitType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TickMarkUnitType", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 GroupingUnit
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GroupingUnit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GroupingUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartAxisUnitTypeEnum GroupingUnitType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartAxisUnitTypeEnum>(this, "GroupingUnitType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "GroupingUnitType", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartGroupingTotalFunctionEnum GroupingTotalFunction
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartGroupingTotalFunctionEnum>(this, "GroupingTotalFunction");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "GroupingTotalFunction", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Orientation
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Orientation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Orientation", value);
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
		public NetOffice.OWC10Api.ChCategoryLabels CategoryLabels
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChCategoryLabels>(this, "CategoryLabels", NetOffice.OWC10Api.ChCategoryLabels.LateBindingApiWrapperType);
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
		/// <param name="value">object value</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Coordinate ValueToPoint(object value)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.Coordinate>(this, "ValueToPoint", NetOffice.OWC10Api.Coordinate.LateBindingApiWrapperType, value);
		}

		#endregion

		#pragma warning restore
	}
}
