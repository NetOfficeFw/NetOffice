using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// Chart
	///</summary>
	public class Chart_ : COMObject
	{
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// <param name="index">optional object Index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_ChartGroups(object index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "ChartGroups", paramsArray);
			ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx
		/// Alias for get_ChartGroups
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object ChartGroups(object index)
		{
			return get_ChartGroups(index);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// <param name="index2">optional object Index2</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_HasAxis(object index1, object index2)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index1, index2);
			object returnItem = Invoker.PropertyGet(this, "HasAxis", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object Index1</param>
        /// <param name="index2">optional object Index2</param>
        /// <param name="value">optional object value</param>
        [SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_HasAxis(object index1, object index2, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index1, index2);
			Invoker.PropertySet(this, "HasAxis", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx
		/// Alias for get_HasAxis
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// <param name="index2">optional object Index2</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object HasAxis(object index1, object index2)
		{
			return get_HasAxis(index1, index2);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_HasAxis(object index1)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index1);
			object returnItem = Invoker.PropertyGet(this, "HasAxis", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object Index1</param>
        /// <param name="value">optional object value</param>
        [SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_HasAxis(object index1, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index1);
			Invoker.PropertySet(this, "HasAxis", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx
		/// Alias for get_HasAxis
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object HasAxis(object index1)
		{
			return get_HasAxis(index1);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// DispatchInterface Chart 
	/// SupportByVersion Word, 14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193446.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Chart : Chart_
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191738.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196350.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool HasTitle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasTitle", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HasTitle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191751.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.ChartTitle ChartTitle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartTitle", paramsArray);
				NetOffice.WordApi.ChartTitle newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ChartTitle.LateBindingApiWrapperType) as NetOffice.WordApi.ChartTitle;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840907.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 DepthPercent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DepthPercent", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DepthPercent", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192611.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 Elevation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Elevation", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Elevation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845244.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 GapDepth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GapDepth", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GapDepth", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836594.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 HeightPercent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HeightPercent", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HeightPercent", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838954.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 Perspective
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Perspective", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Perspective", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838938.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object RightAngleAxes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RightAngleAxes", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RightAngleAxes", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835465.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Rotation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Rotation", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Rotation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196216.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlDisplayBlanksAs DisplayBlanksAs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayBlanksAs", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.XlDisplayBlanksAs)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayBlanksAs", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object ChartGroups
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartGroups", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 SubType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SubType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SubType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Type", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.Corners Corners
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Corners", paramsArray);
				NetOffice.WordApi.Corners newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Corners.LateBindingApiWrapperType) as NetOffice.WordApi.Corners;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836334.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.OfficeApi.Enums.XlChartType ChartType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.XlChartType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ChartType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197158.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool HasDataTable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasDataTable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HasDataTable", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836380.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlRowCol PlotBy
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PlotBy", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.XlRowCol)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PlotBy", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845054.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool HasLegend
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasLegend", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HasLegend", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836685.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Legend Legend
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Legend", paramsArray);
				NetOffice.WordApi.Legend newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Legend.LateBindingApiWrapperType) as NetOffice.WordApi.Legend;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object HasAxis
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasAxis", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HasAxis", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840511.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Walls Walls
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Walls", paramsArray);
				NetOffice.WordApi.Walls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Walls.LateBindingApiWrapperType) as NetOffice.WordApi.Walls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845855.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Floor Floor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Floor", paramsArray);
				NetOffice.WordApi.Floor newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Floor.LateBindingApiWrapperType) as NetOffice.WordApi.Floor;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194655.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.PlotArea PlotArea
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PlotArea", paramsArray);
				NetOffice.WordApi.PlotArea newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.PlotArea.LateBindingApiWrapperType) as NetOffice.WordApi.PlotArea;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196388.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool PlotVisibleOnly
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PlotVisibleOnly", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PlotVisibleOnly", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836658.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.ChartArea ChartArea
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartArea", paramsArray);
				NetOffice.WordApi.ChartArea newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ChartArea.LateBindingApiWrapperType) as NetOffice.WordApi.ChartArea;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823268.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool AutoScaling
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoScaling", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutoScaling", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191967.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.DataTable DataTable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataTable", paramsArray);
				NetOffice.WordApi.DataTable newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.DataTable.LateBindingApiWrapperType) as NetOffice.WordApi.DataTable;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839500.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlBarShape BarShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BarShape", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.XlBarShape)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BarShape", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839285.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Walls SideWall
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SideWall", paramsArray);
				NetOffice.WordApi.Walls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Walls.LateBindingApiWrapperType) as NetOffice.WordApi.Walls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193753.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Walls BackWall
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BackWall", paramsArray);
				NetOffice.WordApi.Walls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Walls.LateBindingApiWrapperType) as NetOffice.WordApi.Walls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195916.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object ChartStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartStyle", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ChartStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836370.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object PivotLayout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PivotLayout", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool HasPivotFields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasPivotFields", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HasPivotFields", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193871.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool ShowDataLabelsOverMaximum
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowDataLabelsOverMaximum", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowDataLabelsOverMaximum", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838941.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.ChartData ChartData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartData", paramsArray);
				NetOffice.WordApi.ChartData newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ChartData.LateBindingApiWrapperType) as NetOffice.WordApi.ChartData;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837462.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Shapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shapes", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835828.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192382.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup Area3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Area3DGroup", paramsArray);
				NetOffice.WordApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.WordApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup Bar3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Bar3DGroup", paramsArray);
				NetOffice.WordApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.WordApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup Column3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Column3DGroup", paramsArray);
				NetOffice.WordApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.WordApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup Line3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Line3DGroup", paramsArray);
				NetOffice.WordApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.WordApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup Pie3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Pie3DGroup", paramsArray);
				NetOffice.WordApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.WordApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ChartGroup SurfaceGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SurfaceGroup", paramsArray);
				NetOffice.WordApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.WordApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198336.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool ShowReportFilterFieldButtons
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowReportFilterFieldButtons", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowReportFilterFieldButtons", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839101.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool ShowLegendFieldButtons
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowLegendFieldButtons", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowLegendFieldButtons", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845234.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool ShowAxisFieldButtons
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowAxisFieldButtons", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowAxisFieldButtons", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834948.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool ShowValueFieldButtons
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowValueFieldButtons", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowValueFieldButtons", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834564.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool ShowAllFieldButtons
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowAllFieldButtons", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowAllFieldButtons", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230485.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
		public NetOffice.WordApi.Enums.XlCategoryLabelLevel CategoryLabelLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CategoryLabelLevel", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.XlCategoryLabelLevel)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CategoryLabelLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232218.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
		public NetOffice.WordApi.Enums.XlSeriesNameLevel SeriesNameLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SeriesNameLevel", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.XlSeriesNameLevel)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SeriesNameLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool HasHiddenContent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasHiddenContent", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231924.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
		public object ChartColor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartColor", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ChartColor", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837270.aspx
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object SeriesCollection(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "SeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837270.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object SeriesCollection()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "SeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		/// <param name="showPercentage">optional object ShowPercentage</param>
		/// <param name="showBubbleSize">optional object ShowBubbleSize</param>
		/// <param name="separator">optional object Separator</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyDataLabels()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyDataLabels(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		/// <param name="showPercentage">optional object ShowPercentage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		/// <param name="showPercentage">optional object ShowPercentage</param>
		/// <param name="showBubbleSize">optional object ShowBubbleSize</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType ChartType</param>
		/// <param name="typeName">optional object TypeName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chartType, typeName);
			Invoker.Method(this, "ApplyCustomType", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType ChartType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chartType);
			Invoker.Method(this, "ApplyCustomType", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839889.aspx
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="elementID">Int32 ElementID</param>
		/// <param name="arg1">Int32 Arg1</param>
		/// <param name="arg2">Int32 Arg2</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822921.aspx
		/// </summary>
		/// <param name="source">string Source</param>
		/// <param name="plotBy">optional object PlotBy</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SetSourceData(string source, object plotBy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, plotBy);
			Invoker.Method(this, "SetSourceData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822921.aspx
		/// </summary>
		/// <param name="source">string Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SetSourceData(string source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			Invoker.Method(this, "SetSourceData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		/// <param name="axisGroup">optional NetOffice.WordApi.Enums.XlAxisGroup AxisGroup = 1</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Axes(object type, object axisGroup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, axisGroup);
			object returnItem = Invoker.MethodReturn(this, "Axes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Axes()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Axes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Axes(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "Axes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="gallery">Int32 Gallery</param>
		/// <param name="format">optional object Format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void AutoFormat(Int32 gallery, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gallery, format);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="gallery">Int32 Gallery</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void AutoFormat(Int32 gallery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gallery);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197864.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SetBackgroundPicture(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SetBackgroundPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		/// <param name="seriesLabels">optional object SeriesLabels</param>
		/// <param name="hasLegend">optional object HasLegend</param>
		/// <param name="title">optional object Title</param>
		/// <param name="categoryTitle">optional object CategoryTitle</param>
		/// <param name="valueTitle">optional object ValueTitle</param>
		/// <param name="extraTitle">optional object ExtraTitle</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle, object extraTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle, extraTitle);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard(object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		/// <param name="seriesLabels">optional object SeriesLabels</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		/// <param name="seriesLabels">optional object SeriesLabels</param>
		/// <param name="hasLegend">optional object HasLegend</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		/// <param name="seriesLabels">optional object SeriesLabels</param>
		/// <param name="hasLegend">optional object HasLegend</param>
		/// <param name="title">optional object Title</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		/// <param name="seriesLabels">optional object SeriesLabels</param>
		/// <param name="hasLegend">optional object HasLegend</param>
		/// <param name="title">optional object Title</param>
		/// <param name="categoryTitle">optional object CategoryTitle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		/// <param name="seriesLabels">optional object SeriesLabels</param>
		/// <param name="hasLegend">optional object HasLegend</param>
		/// <param name="title">optional object Title</param>
		/// <param name="categoryTitle">optional object CategoryTitle</param>
		/// <param name="valueTitle">optional object ValueTitle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx
		/// </summary>
		/// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="format">optional NetOffice.WordApi.Enums.XlCopyPictureFormat Format = -4147</param>
		/// <param name="size">optional NetOffice.WordApi.Enums.XlPictureAppearance Size = 2</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void CopyPicture(object appearance, object format, object size)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(appearance, format, size);
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void CopyPicture()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx
		/// </summary>
		/// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void CopyPicture(object appearance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(appearance);
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx
		/// </summary>
		/// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="format">optional NetOffice.WordApi.Enums.XlCopyPictureFormat Format = -4147</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void CopyPicture(object appearance, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(appearance, format);
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839410.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void Paste(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839410.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="interactive">optional object Interactive</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool Export(string fileName, object filterName, object interactive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, filterName, interactive);
			object returnItem = Invoker.MethodReturn(this, "Export", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool Export(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "Export", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="filterName">optional object FilterName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool Export(string fileName, object filterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, filterName);
			object returnItem = Invoker.MethodReturn(this, "Export", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839841.aspx
		/// </summary>
		/// <param name="name">object Name</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SetDefaultChart(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "SetDefaultChart", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845631.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyChartTemplate(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "ApplyChartTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839083.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SaveChartTemplate(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SaveChartTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193390.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ClearToMatchStyle()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearToMatchStyle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840397.aspx
		/// </summary>
		/// <param name="layout">Int32 Layout</param>
		/// <param name="chartType">optional object ChartType</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyLayout(Int32 layout, object chartType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout, chartType);
			Invoker.Method(this, "ApplyLayout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840397.aspx
		/// </summary>
		/// <param name="layout">Int32 Layout</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ApplyLayout(Int32 layout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout);
			Invoker.Method(this, "ApplyLayout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192135.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void Refresh()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Refresh", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838552.aspx
		/// </summary>
		/// <param name="element">NetOffice.OfficeApi.Enums.MsoChartElementType Element</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType element)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(element);
			Invoker.Method(this, "SetElement", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object AreaGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "AreaGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object AreaGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AreaGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object BarGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "BarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object BarGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "BarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object ColumnGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "ColumnGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object ColumnGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ColumnGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object LineGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "LineGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object LineGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "LineGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object PieGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "PieGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object PieGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PieGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object DoughnutGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "DoughnutGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object DoughnutGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DoughnutGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object RadarGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "RadarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object RadarGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "RadarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object XYGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "XYGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object XYGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "XYGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840074.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Delete()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Delete", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx
		/// </summary>
		/// <param name="before">optional object Before</param>
		/// <param name="after">optional object After</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void Copy(object before, object after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(before, after);
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx
		/// </summary>
		/// <param name="before">optional object Before</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void Copy(object before)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(before);
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191928.aspx
		/// </summary>
		/// <param name="replace">optional object Replace</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Select(object replace)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(replace);
			object returnItem = Invoker.MethodReturn(this, "Select", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191928.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Select()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Select", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229848.aspx
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Word", 15, 16)]
		public object FullSeriesCollection(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "FullSeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229848.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 15, 16)]
		public object FullSeriesCollection()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FullSeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 15, 16)]
		public void DeleteHiddenContent()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DeleteHiddenContent", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230203.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
		public void ClearToMatchColorStyle()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearToMatchColorStyle", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}