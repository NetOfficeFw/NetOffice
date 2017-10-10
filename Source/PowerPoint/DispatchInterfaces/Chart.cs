using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// Chart
	/// </summary>
	[SyntaxBypass]
 	public class Chart_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Chart_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Chart_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index1">optional object index1</param>
		/// <param name="index2">optional object index2</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_HasAxis(object index1, object index2)
		{
			return Factory.ExecuteVariantPropertyGet(this, "HasAxis", index1, index2);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index1">optional object index1</param>
		/// <param name="index2">optional object index2</param>
        /// <param name="value">optional object value</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_HasAxis(object index1, object index2, object value)
		{
			Factory.ExecutePropertySet(this, "HasAxis", index1, index2, value);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Alias for get_HasAxis
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx </remarks>
		/// <param name="index1">optional object index1</param>
		/// <param name="index2">optional object index2</param>
		[SupportByVersion("PowerPoint", 14,15,16), Redirect("get_HasAxis")]
		public object HasAxis(object index1, object index2)
		{
			return get_HasAxis(index1, index2);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index1">optional object index1</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_HasAxis(object index1)
		{
			return Factory.ExecuteVariantPropertyGet(this, "HasAxis", index1);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index1">optional object index1</param>
        /// <param name="value">optional object value</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_HasAxis(object index1, object value)
		{
			Factory.ExecutePropertySet(this, "HasAxis", index1, value);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Alias for get_HasAxis
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx </remarks>
		/// <param name="index1">optional object index1</param>
		[SupportByVersion("PowerPoint", 14,15,16), Redirect("get_HasAxis")]
		public object HasAxis(object index1)
		{
			return get_HasAxis(index1);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface Chart 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744663.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Chart : Chart_
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
                    _type = typeof(Chart);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Chart(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Chart(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746116.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744954.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745140.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool HasDataTable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasDataTable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasDataTable", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744071.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlRowCol PlotBy
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlRowCol>(this, "PlotBy");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PlotBy", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746809.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.DataTable DataTable
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DataTable>(this, "DataTable", NetOffice.PowerPointApi.DataTable.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746790.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlBarShape BarShape
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlBarShape>(this, "BarShape");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BarShape", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744381.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Walls SideWall
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Walls>(this, "SideWall", NetOffice.PowerPointApi.Walls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744079.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Walls BackWall
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Walls>(this, "BackWall", NetOffice.PowerPointApi.Walls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743954.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ChartStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ChartStyle");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ChartStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool HasPivotFields
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasPivotFields");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasPivotFields", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745647.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ShowDataLabelsOverMaximum
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowDataLabelsOverMaximum");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowDataLabelsOverMaximum", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744089.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.ChartData ChartData
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartData>(this, "ChartData", NetOffice.PowerPointApi.ChartData.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746059.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shapes Shapes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shapes>(this, "Shapes", NetOffice.PowerPointApi.Shapes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746336.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.ChartGroup Area3DGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "Area3DGroup", NetOffice.PowerPointApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.ChartGroup Bar3DGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "Bar3DGroup", NetOffice.PowerPointApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.ChartGroup Column3DGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "Column3DGroup", NetOffice.PowerPointApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.ChartGroup Line3DGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "Line3DGroup", NetOffice.PowerPointApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.ChartGroup Pie3DGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "Pie3DGroup", NetOffice.PowerPointApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.ChartGroup SurfaceGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "SurfaceGroup", NetOffice.PowerPointApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745066.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744513.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool AutoScaling
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoScaling");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoScaling", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744327.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.ChartArea ChartArea
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartArea>(this, "ChartArea", NetOffice.PowerPointApi.ChartArea.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743961.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.ChartTitle ChartTitle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartTitle>(this, "ChartTitle", NetOffice.PowerPointApi.ChartTitle.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.Corners Corners
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Corners>(this, "Corners", NetOffice.PowerPointApi.Corners.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746755.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 DepthPercent
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DepthPercent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DepthPercent", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745600.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlDisplayBlanksAs DisplayBlanksAs
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlDisplayBlanksAs>(this, "DisplayBlanksAs");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayBlanksAs", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745750.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Elevation
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Elevation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Elevation", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745846.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Floor Floor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Floor>(this, "Floor", NetOffice.PowerPointApi.Floor.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746511.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object HasAxis
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HasAxis");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "HasAxis", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743935.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746534.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745241.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 HeightPercent
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HeightPercent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HeightPercent", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744151.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Legend Legend
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Legend>(this, "Legend", NetOffice.PowerPointApi.Legend.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744105.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743957.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746093.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.PlotArea PlotArea
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PlotArea>(this, "PlotArea", NetOffice.PowerPointApi.PlotArea.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745749.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool PlotVisibleOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PlotVisibleOnly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PlotVisibleOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744814.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object RightAngleAxes
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RightAngleAxes");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "RightAngleAxes", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745024.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Rotation
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Rotation");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Rotation", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 Subtype
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Subtype");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Subtype", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746542.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Walls Walls
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Walls>(this, "Walls", NetOffice.PowerPointApi.Walls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745294.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.ChartFormat Format
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartFormat>(this, "Format", NetOffice.PowerPointApi.ChartFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745821.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ShowReportFilterFieldButtons
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowReportFilterFieldButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowReportFilterFieldButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743877.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ShowLegendFieldButtons
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowLegendFieldButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowLegendFieldButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744539.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ShowAxisFieldButtons
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAxisFieldButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAxisFieldButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746204.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ShowValueFieldButtons
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowValueFieldButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowValueFieldButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744868.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ShowAllFieldButtons
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAllFieldButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAllFieldButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746125.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string AlternativeText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AlternativeText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlternativeText", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745833.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string Title
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Title");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Title", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229264.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Enums.XlCategoryLabelLevel CategoryLabelLevel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlCategoryLabelLevel>(this, "CategoryLabelLevel");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CategoryLabelLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228519.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Enums.XlSeriesNameLevel SeriesNameLevel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlSeriesNameLevel>(this, "SeriesNameLevel");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SeriesNameLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 15, 16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool HasHiddenContent
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasHiddenContent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230443.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public object ChartColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ChartColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ChartColor", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		/// <param name="showBubbleSize">optional object showBubbleSize</param>
		/// <param name="separator">optional object separator</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyDataLabels()
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyDataLabels(object type)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", type, legendKey);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", type, legendKey, autoText);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", type, legendKey, autoText, hasLeaderLines);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		/// <param name="showBubbleSize">optional object showBubbleSize</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
		/// <param name="typeName">optional object typeName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName)
		{
			 Factory.ExecuteMethod(this, "ApplyCustomType", chartType, typeName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
		{
			 Factory.ExecuteMethod(this, "ApplyCustomType", chartType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746151.aspx </remarks>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="elementID">Int32 elementID</param>
		/// <param name="arg1">Int32 arg1</param>
		/// <param name="arg2">Int32 arg2</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void GetChartElement(Int32 x, Int32 y, Int32 elementID, Int32 arg1, Int32 arg2)
		{
			 Factory.ExecuteMethod(this, "GetChartElement", new object[]{ x, y, elementID, arg1, arg2 });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746759.aspx </remarks>
		/// <param name="source">string source</param>
		/// <param name="plotBy">optional object plotBy</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void SetSourceData(string source, object plotBy)
		{
			 Factory.ExecuteMethod(this, "SetSourceData", source, plotBy);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746759.aspx </remarks>
		/// <param name="source">string source</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void SetSourceData(string source)
		{
			 Factory.ExecuteMethod(this, "SetSourceData", source);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="gallery">Int32 gallery</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void AutoFormat(Int32 gallery, object format)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", gallery, format);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="gallery">Int32 gallery</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void AutoFormat(Int32 gallery)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", gallery);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745424.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void SetBackgroundPicture(string fileName)
		{
			 Factory.ExecuteMethod(this, "SetBackgroundPicture", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746056.aspx </remarks>
		/// <param name="type">optional object type</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Paste(object type)
		{
			 Factory.ExecuteMethod(this, "Paste", type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746056.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745864.aspx </remarks>
		/// <param name="name">object name</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void SetDefaultChart(object name)
		{
			 Factory.ExecuteMethod(this, "SetDefaultChart", name);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744899.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyChartTemplate(string fileName)
		{
			 Factory.ExecuteMethod(this, "ApplyChartTemplate", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744919.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void SaveChartTemplate(string fileName)
		{
			 Factory.ExecuteMethod(this, "SaveChartTemplate", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746785.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ClearToMatchStyle()
		{
			 Factory.ExecuteMethod(this, "ClearToMatchStyle");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745663.aspx </remarks>
		/// <param name="layout">Int32 layout</param>
		/// <param name="chartType">optional object chartType</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyLayout(Int32 layout, object chartType)
		{
			 Factory.ExecuteMethod(this, "ApplyLayout", layout, chartType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745663.aspx </remarks>
		/// <param name="layout">Int32 layout</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyLayout(Int32 layout)
		{
			 Factory.ExecuteMethod(this, "ApplyLayout", layout);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745006.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object AreaGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "AreaGroups", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object AreaGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "AreaGroups");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object BarGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "BarGroups", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object BarGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "BarGroups");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ColumnGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "ColumnGroups", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ColumnGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "ColumnGroups");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object LineGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "LineGroups", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object LineGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "LineGroups");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object PieGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "PieGroups", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object PieGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "PieGroups");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object DoughnutGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "DoughnutGroups", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object DoughnutGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "DoughnutGroups");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object RadarGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "RadarGroups", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object RadarGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "RadarGroups");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object XYGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "XYGroups", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object XYGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "XYGroups");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void _ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
		{
			 Factory.ExecuteMethod(this, "_ApplyDataLabels", type, legendKey, autoText, hasLeaderLines);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void _ApplyDataLabels()
		{
			 Factory.ExecuteMethod(this, "_ApplyDataLabels");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void _ApplyDataLabels(object type)
		{
			 Factory.ExecuteMethod(this, "_ApplyDataLabels", type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void _ApplyDataLabels(object type, object legendKey)
		{
			 Factory.ExecuteMethod(this, "_ApplyDataLabels", type, legendKey);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void _ApplyDataLabels(object type, object legendKey, object autoText)
		{
			 Factory.ExecuteMethod(this, "_ApplyDataLabels", type, legendKey, autoText);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745216.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="axisGroup">optional NetOffice.PowerPointApi.Enums.XlAxisGroup AxisGroup = 1</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Axes(object type, object axisGroup)
		{
			return Factory.ExecuteVariantMethodGet(this, "Axes", type, axisGroup);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745216.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Axes()
		{
			return Factory.ExecuteVariantMethodGet(this, "Axes");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745216.aspx </remarks>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Axes(object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "Axes", type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744238.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ChartGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "ChartGroups", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744238.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ChartGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "ChartGroups");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="hasLegend">optional object hasLegend</param>
		/// <param name="title">optional object title</param>
		/// <param name="categoryTitle">optional object categoryTitle</param>
		/// <param name="valueTitle">optional object valueTitle</param>
		/// <param name="extraTitle">optional object extraTitle</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle, object extraTitle)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle, extraTitle });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard()
		{
			 Factory.ExecuteMethod(this, "ChartWizard");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard(object source)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", source);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard(object source, object gallery)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", source, gallery);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", source, gallery, format);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", source, gallery, format, plotBy);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="hasLegend">optional object hasLegend</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="hasLegend">optional object hasLegend</param>
		/// <param name="title">optional object title</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="hasLegend">optional object hasLegend</param>
		/// <param name="title">optional object title</param>
		/// <param name="categoryTitle">optional object categoryTitle</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="hasLegend">optional object hasLegend</param>
		/// <param name="title">optional object title</param>
		/// <param name="categoryTitle">optional object categoryTitle</param>
		/// <param name="valueTitle">optional object valueTitle</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745934.aspx </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Copy(object before, object after)
		{
			 Factory.ExecuteMethod(this, "Copy", before, after);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745934.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745934.aspx </remarks>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Copy(object before)
		{
			 Factory.ExecuteMethod(this, "Copy", before);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
		/// <param name="appearance">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="format">optional NetOffice.PowerPointApi.Enums.XlCopyPictureFormat Format = -4147</param>
		/// <param name="size">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Size = 2</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CopyPicture(object appearance, object format, object size)
		{
			 Factory.ExecuteMethod(this, "CopyPicture", appearance, format, size);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CopyPicture()
		{
			 Factory.ExecuteMethod(this, "CopyPicture");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
		/// <param name="appearance">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Appearance = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CopyPicture(object appearance)
		{
			 Factory.ExecuteMethod(this, "CopyPicture", appearance);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
		/// <param name="appearance">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="format">optional NetOffice.PowerPointApi.Enums.XlCopyPictureFormat Format = -4147</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CopyPicture(object appearance, object format)
		{
			 Factory.ExecuteMethod(this, "CopyPicture", appearance, format);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745109.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744128.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="interactive">optional object interactive</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool Export(string fileName, object filterName, object interactive)
		{
			return Factory.ExecuteBoolMethodGet(this, "Export", fileName, filterName, interactive);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744128.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool Export(string fileName)
		{
			return Factory.ExecuteBoolMethodGet(this, "Export", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744128.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="filterName">optional object filterName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool Export(string fileName, object filterName)
		{
			return Factory.ExecuteBoolMethodGet(this, "Export", fileName, filterName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745013.aspx </remarks>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Select(object replace)
		{
			 Factory.ExecuteMethod(this, "Select", replace);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745013.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745538.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object SeriesCollection(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "SeriesCollection", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745538.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object SeriesCollection()
		{
			return Factory.ExecuteVariantMethodGet(this, "SeriesCollection");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746262.aspx </remarks>
		/// <param name="element">NetOffice.OfficeApi.Enums.MsoChartElementType element</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType element)
		{
			 Factory.ExecuteMethod(this, "SetElement", element);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228028.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		public object FullSeriesCollection(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "FullSeriesCollection", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228028.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public object FullSeriesCollection()
		{
			return Factory.ExecuteVariantMethodGet(this, "FullSeriesCollection");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void DeleteHiddenContent()
		{
			 Factory.ExecuteMethod(this, "DeleteHiddenContent");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227229.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ClearToMatchColorStyle()
		{
			 Factory.ExecuteMethod(this, "ClearToMatchColorStyle");
		}

		#endregion

		#pragma warning restore
	}
}
