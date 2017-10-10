using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_ChartGroups(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "ChartGroups", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx
		/// Alias for get_ChartGroups
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Word", 14,15,16), ProxyResult, Redirect("get_ChartGroups")]
		public object ChartGroups(object index)
		{
			return get_ChartGroups(index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index1">optional object index1</param>
		/// <param name="index2">optional object index2</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_HasAxis(object index1, object index2)
		{
			return Factory.ExecuteVariantPropertyGet(this, "HasAxis", index1, index2);
		}

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_HasAxis(object index1, object index2, object value)
		{
			Factory.ExecutePropertySet(this, "HasAxis", index1, index2, value);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Alias for get_HasAxis
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx </remarks>
		/// <param name="index1">optional object index1</param>
		/// <param name="index2">optional object index2</param>
		[SupportByVersion("Word", 14,15,16), Redirect("get_HasAxis")]
		public object HasAxis(object index1, object index2)
		{
			return get_HasAxis(index1, index2);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index1">optional object index1</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_HasAxis(object index1)
		{
			return Factory.ExecuteVariantPropertyGet(this, "HasAxis", index1);
		}

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_HasAxis(object index1, object value)
		{
			Factory.ExecutePropertySet(this, "HasAxis", index1, value);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Alias for get_HasAxis
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx </remarks>
		/// <param name="index1">optional object index1</param>
		[SupportByVersion("Word", 14,15,16), Redirect("get_HasAxis")]
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
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193446.aspx </remarks>
	[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191738.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196350.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191751.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ChartTitle ChartTitle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartTitle>(this, "ChartTitle", NetOffice.WordApi.ChartTitle.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840907.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192611.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845244.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836594.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838954.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838938.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835465.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196216.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlDisplayBlanksAs DisplayBlanksAs
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlDisplayBlanksAs>(this, "DisplayBlanksAs");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayBlanksAs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		public object ChartGroups
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ChartGroups");
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
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.Corners Corners
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Corners>(this, "Corners", NetOffice.WordApi.Corners.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836334.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197158.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836380.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlRowCol PlotBy
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlRowCol>(this, "PlotBy");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PlotBy", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845054.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836685.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Legend Legend
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Legend>(this, "Legend", NetOffice.WordApi.Legend.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840511.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Walls Walls
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Walls>(this, "Walls", NetOffice.WordApi.Walls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845855.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Floor Floor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Floor>(this, "Floor", NetOffice.WordApi.Floor.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194655.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.PlotArea PlotArea
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.PlotArea>(this, "PlotArea", NetOffice.WordApi.PlotArea.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196388.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836658.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ChartArea ChartArea
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartArea>(this, "ChartArea", NetOffice.WordApi.ChartArea.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823268.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191967.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.DataTable DataTable
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.DataTable>(this, "DataTable", NetOffice.WordApi.DataTable.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839500.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839285.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Walls SideWall
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Walls>(this, "SideWall", NetOffice.WordApi.Walls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193753.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Walls BackWall
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Walls>(this, "BackWall", NetOffice.WordApi.Walls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195916.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836370.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		public object PivotLayout
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "PivotLayout");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193871.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838941.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ChartData ChartData
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartData>(this, "ChartData", NetOffice.WordApi.ChartData.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837462.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		public object Shapes
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Shapes");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835828.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192382.aspx </remarks>
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
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup Area3DGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "Area3DGroup", NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup Bar3DGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "Bar3DGroup", NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup Column3DGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "Column3DGroup", NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup Line3DGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "Line3DGroup", NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup Pie3DGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "Pie3DGroup", NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup SurfaceGroup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "SurfaceGroup", NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198336.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839101.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845234.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834948.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834564.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230485.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Enums.XlCategoryLabelLevel CategoryLabelLevel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlCategoryLabelLevel>(this, "CategoryLabelLevel");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CategoryLabelLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232218.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Enums.XlSeriesNameLevel SeriesNameLevel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlSeriesNameLevel>(this, "SeriesNameLevel");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SeriesNameLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 15, 16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool HasHiddenContent
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasHiddenContent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231924.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837270.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Word", 14,15,16)]
		public object SeriesCollection(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "SeriesCollection", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837270.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object SeriesCollection()
		{
			return Factory.ExecuteVariantMethodGet(this, "SeriesCollection");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
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
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyDataLabels()
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyDataLabels(object type)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", type);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", type, legendKey);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", type, legendKey, autoText);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", type, legendKey, autoText, hasLeaderLines);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
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
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
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
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
		{
			 Factory.ExecuteMethod(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
		/// <param name="typeName">optional object typeName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName)
		{
			 Factory.ExecuteMethod(this, "ApplyCustomType", chartType, typeName);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
		{
			 Factory.ExecuteMethod(this, "ApplyCustomType", chartType);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839889.aspx </remarks>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="elementID">Int32 elementID</param>
		/// <param name="arg1">Int32 arg1</param>
		/// <param name="arg2">Int32 arg2</param>
		[SupportByVersion("Word", 14,15,16)]
		public void GetChartElement(Int32 x, Int32 y, out Int32 elementID, out Int32 arg1, out Int32 arg2)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true,true);
			elementID = 0;
			arg1 = 0;
			arg2 = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, elementID, arg1, arg2);
			Invoker.Method(this, "GetChartElement", paramsArray, modifiers);
			elementID = (Int32)paramsArray[2];
			arg1 = (Int32)paramsArray[3];
			arg2 = (Int32)paramsArray[4];
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822921.aspx </remarks>
		/// <param name="source">string source</param>
		/// <param name="plotBy">optional object plotBy</param>
		[SupportByVersion("Word", 14,15,16)]
		public void SetSourceData(string source, object plotBy)
		{
			 Factory.ExecuteMethod(this, "SetSourceData", source, plotBy);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822921.aspx </remarks>
		/// <param name="source">string source</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SetSourceData(string source)
		{
			 Factory.ExecuteMethod(this, "SetSourceData", source);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="axisGroup">optional NetOffice.WordApi.Enums.XlAxisGroup AxisGroup = 1</param>
		[SupportByVersion("Word", 14,15,16)]
		public object Axes(object type, object axisGroup)
		{
			return Factory.ExecuteVariantMethodGet(this, "Axes", type, axisGroup);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object Axes()
		{
			return Factory.ExecuteVariantMethodGet(this, "Axes");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx </remarks>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object Axes(object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "Axes", type);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="gallery">Int32 gallery</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public void AutoFormat(Int32 gallery, object format)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", gallery, format);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="gallery">Int32 gallery</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void AutoFormat(Int32 gallery)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", gallery);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197864.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 14,15,16)]
		public void SetBackgroundPicture(string fileName)
		{
			 Factory.ExecuteMethod(this, "SetBackgroundPicture", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
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
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle, object extraTitle)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle, extraTitle });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard()
		{
			 Factory.ExecuteMethod(this, "ChartWizard");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard(object source)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", source);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", source, gallery);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", source, gallery, format);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", source, gallery, format, plotBy);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="hasLegend">optional object hasLegend</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
		/// <param name="source">optional object source</param>
		/// <param name="gallery">optional object gallery</param>
		/// <param name="format">optional object format</param>
		/// <param name="plotBy">optional object plotBy</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="hasLegend">optional object hasLegend</param>
		/// <param name="title">optional object title</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
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
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
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
		[SupportByVersion("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle)
		{
			 Factory.ExecuteMethod(this, "ChartWizard", new object[]{ source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
		/// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="format">optional NetOffice.WordApi.Enums.XlCopyPictureFormat Format = -4147</param>
		/// <param name="size">optional NetOffice.WordApi.Enums.XlPictureAppearance Size = 2</param>
		[SupportByVersion("Word", 14,15,16)]
		public void CopyPicture(object appearance, object format, object size)
		{
			 Factory.ExecuteMethod(this, "CopyPicture", appearance, format, size);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void CopyPicture()
		{
			 Factory.ExecuteMethod(this, "CopyPicture");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
		/// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void CopyPicture(object appearance)
		{
			 Factory.ExecuteMethod(this, "CopyPicture", appearance);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
		/// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="format">optional NetOffice.WordApi.Enums.XlCopyPictureFormat Format = -4147</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void CopyPicture(object appearance, object format)
		{
			 Factory.ExecuteMethod(this, "CopyPicture", appearance, format);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839410.aspx </remarks>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Word", 14,15,16)]
		public void Paste(object type)
		{
			 Factory.ExecuteMethod(this, "Paste", type);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839410.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="interactive">optional object interactive</param>
		[SupportByVersion("Word", 14,15,16)]
		public bool Export(string fileName, object filterName, object interactive)
		{
			return Factory.ExecuteBoolMethodGet(this, "Export", fileName, filterName, interactive);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public bool Export(string fileName)
		{
			return Factory.ExecuteBoolMethodGet(this, "Export", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="filterName">optional object filterName</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public bool Export(string fileName, object filterName)
		{
			return Factory.ExecuteBoolMethodGet(this, "Export", fileName, filterName);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839841.aspx </remarks>
		/// <param name="name">object name</param>
		[SupportByVersion("Word", 14,15,16)]
		public void SetDefaultChart(object name)
		{
			 Factory.ExecuteMethod(this, "SetDefaultChart", name);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845631.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyChartTemplate(string fileName)
		{
			 Factory.ExecuteMethod(this, "ApplyChartTemplate", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839083.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 14,15,16)]
		public void SaveChartTemplate(string fileName)
		{
			 Factory.ExecuteMethod(this, "SaveChartTemplate", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193390.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public void ClearToMatchStyle()
		{
			 Factory.ExecuteMethod(this, "ClearToMatchStyle");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840397.aspx </remarks>
		/// <param name="layout">Int32 layout</param>
		/// <param name="chartType">optional object chartType</param>
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyLayout(Int32 layout, object chartType)
		{
			 Factory.ExecuteMethod(this, "ApplyLayout", layout, chartType);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840397.aspx </remarks>
		/// <param name="layout">Int32 layout</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyLayout(Int32 layout)
		{
			 Factory.ExecuteMethod(this, "ApplyLayout", layout);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192135.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838552.aspx </remarks>
		/// <param name="element">NetOffice.OfficeApi.Enums.MsoChartElementType element</param>
		[SupportByVersion("Word", 14,15,16)]
		public void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType element)
		{
			 Factory.ExecuteMethod(this, "SetElement", element);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public object AreaGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "AreaGroups", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object AreaGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "AreaGroups");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public object BarGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "BarGroups", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object BarGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "BarGroups");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public object ColumnGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "ColumnGroups", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object ColumnGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "ColumnGroups");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public object LineGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "LineGroups", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object LineGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "LineGroups");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public object PieGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "PieGroups", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object PieGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "PieGroups");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public object DoughnutGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "DoughnutGroups", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object DoughnutGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "DoughnutGroups");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public object RadarGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "RadarGroups", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object RadarGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "RadarGroups");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public object XYGroups(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "XYGroups", index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object XYGroups()
		{
			return Factory.ExecuteVariantMethodGet(this, "XYGroups");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840074.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public object Delete()
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("Word", 14,15,16)]
		public void Copy(object before, object after)
		{
			 Factory.ExecuteMethod(this, "Copy", before, after);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx </remarks>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void Copy(object before)
		{
			 Factory.ExecuteMethod(this, "Copy", before);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191928.aspx </remarks>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("Word", 14,15,16)]
		public object Select(object replace)
		{
			return Factory.ExecuteVariantMethodGet(this, "Select", replace);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191928.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object Select()
		{
			return Factory.ExecuteVariantMethodGet(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229848.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Word", 15, 16)]
		public object FullSeriesCollection(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "FullSeriesCollection", index);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229848.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public object FullSeriesCollection()
		{
			return Factory.ExecuteVariantMethodGet(this, "FullSeriesCollection");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 15, 16)]
		public void DeleteHiddenContent()
		{
			 Factory.ExecuteMethod(this, "DeleteHiddenContent");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230203.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public void ClearToMatchColorStyle()
		{
			 Factory.ExecuteMethod(this, "ClearToMatchColorStyle");
		}

		#endregion

		#pragma warning restore
	}
}
