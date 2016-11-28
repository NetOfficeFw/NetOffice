using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// IMsoChart
	///</summary>
	public class IMsoChart_ : COMObject
	{
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMsoChart_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart_() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="pvarIndex">optional object pvarIndex</param>
		/// <param name="varIgallery">optional object varIgallery</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_ChartGroups(object pvarIndex, object varIgallery)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(pvarIndex, varIgallery);
			object returnItem = Invoker.PropertyGet(this, "ChartGroups", paramsArray);
			ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_ChartGroups
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="pvarIndex">optional object pvarIndex</param>
		/// <param name="varIgallery">optional object varIgallery</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object ChartGroups(object pvarIndex, object varIgallery)
		{
			return get_ChartGroups(pvarIndex, varIgallery);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="pvarIndex">optional object pvarIndex</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_ChartGroups(object pvarIndex)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(pvarIndex);
			object returnItem = Invoker.PropertyGet(this, "ChartGroups", paramsArray);
			ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_ChartGroups
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="pvarIndex">optional object pvarIndex</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object ChartGroups(object pvarIndex)
		{
			return get_ChartGroups(pvarIndex);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="axisType">optional object axisType</param>
		/// <param name="axisGroup">optional object AxisGroup</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_HasAxis(object axisType, object axisGroup)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(axisType, axisGroup);
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object AxisGroup</param>
        /// <param name="value">optional object value</param>
        [SupportByVersionAttribute("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_HasAxis(object axisType, object axisGroup, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(axisType, axisGroup);
			Invoker.PropertySet(this, "HasAxis", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_HasAxis
		/// </summary>
		/// <param name="axisType">optional object axisType</param>
		/// <param name="axisGroup">optional object AxisGroup</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object HasAxis(object axisType, object axisGroup)
		{
			return get_HasAxis(axisType, axisGroup);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="axisType">optional object axisType</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_HasAxis(object axisType)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(axisType);
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersionAttribute("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_HasAxis(object axisType, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(axisType);
			Invoker.PropertySet(this, "HasAxis", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_HasAxis
		/// </summary>
		/// <param name="axisType">optional object axisType</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object HasAxis(object axisType)
		{
			return get_HasAxis(axisType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="fBackWall">optional bool fBackWall</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.IMsoWalls get_Walls(object fBackWall)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(fBackWall);
			object returnItem = Invoker.PropertyGet(this, "Walls", paramsArray);
			NetOffice.OfficeApi.IMsoWalls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoWalls.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoWalls;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Walls
		/// </summary>
		/// <param name="fBackWall">optional bool fBackWall</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoWalls Walls(object fBackWall)
		{
			return get_Walls(fBackWall);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// DispatchInterface IMsoChart 
	/// SupportByVersion Office, 12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IMsoChart : IMsoChart_
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
                    _type = typeof(IMsoChart);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMsoChart(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoChart(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoChartTitle ChartTitle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartTitle", paramsArray);
				NetOffice.OfficeApi.IMsoChartTitle newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OfficeApi.IMsoChartTitle;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlDisplayBlanksAs DisplayBlanksAs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayBlanksAs", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.XlDisplayBlanksAs)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayBlanksAs", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool ProtectData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectData", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ProtectData", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool ProtectFormatting
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectFormatting", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ProtectFormatting", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool ProtectGoalSeek
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectGoalSeek", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ProtectGoalSeek", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool ProtectSelection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectSelection", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ProtectSelection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool ProtectChartObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectChartObjects", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ProtectChartObjects", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.IMsoCorners Corners
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Corners", paramsArray);
				NetOffice.OfficeApi.IMsoCorners newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoCorners.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoCorners;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlRowCol PlotBy
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PlotBy", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.XlRowCol)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PlotBy", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoLegend Legend
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Legend", paramsArray);
				NetOffice.OfficeApi.IMsoLegend newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoLegend.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoLegend;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoWalls Walls
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Walls", paramsArray);
				NetOffice.OfficeApi.IMsoWalls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoWalls.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoWalls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoFloor Floor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Floor", paramsArray);
				NetOffice.OfficeApi.IMsoFloor newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoFloor.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoFloor;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoPlotArea PlotArea
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PlotArea", paramsArray);
				NetOffice.OfficeApi.IMsoPlotArea newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoPlotArea.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoPlotArea;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoChartArea ChartArea
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartArea", paramsArray);
				NetOffice.OfficeApi.IMsoChartArea newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoChartArea.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoChartArea;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoDataTable DataTable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataTable", paramsArray);
				NetOffice.OfficeApi.IMsoDataTable newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoDataTable.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoDataTable;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.XlBarShape BarShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BarShape", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.XlBarShape)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BarShape", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoWalls SideWall
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SideWall", paramsArray);
				NetOffice.OfficeApi.IMsoWalls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoWalls.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoWalls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoWalls BackWall
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BackWall", paramsArray);
				NetOffice.OfficeApi.IMsoWalls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoWalls.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoWalls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object Selection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Selection", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoChartData ChartData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartData", paramsArray);
				NetOffice.OfficeApi.IMsoChartData newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoChartData.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoChartData;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoChartFormat Format
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Format", paramsArray);
				NetOffice.OfficeApi.IMsoChartFormat newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoChartFormat.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoChartFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shapes Shapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shapes", paramsArray);
				NetOffice.OfficeApi.Shapes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Shapes.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shapes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.IMsoChartGroup Area3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Area3DGroup", paramsArray);
				NetOffice.OfficeApi.IMsoChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoChartGroup.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.IMsoChartGroup Bar3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Bar3DGroup", paramsArray);
				NetOffice.OfficeApi.IMsoChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoChartGroup.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.IMsoChartGroup Column3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Column3DGroup", paramsArray);
				NetOffice.OfficeApi.IMsoChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoChartGroup.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.IMsoChartGroup Line3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Line3DGroup", paramsArray);
				NetOffice.OfficeApi.IMsoChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoChartGroup.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.IMsoChartGroup Pie3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Pie3DGroup", paramsArray);
				NetOffice.OfficeApi.IMsoChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoChartGroup.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.IMsoChartGroup SurfaceGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SurfaceGroup", paramsArray);
				NetOffice.OfficeApi.IMsoChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoChartGroup.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
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
		/// SupportByVersion Office 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 15, 16)]
		public bool ProtectChartSheetFormatting
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectChartSheetFormatting", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ProtectChartSheetFormatting", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 15, 16)]
		public NetOffice.OfficeApi.Enums.XlCategoryLabelLevel CategoryLabelLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CategoryLabelLevel", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.XlCategoryLabelLevel)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CategoryLabelLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 15, 16)]
		public NetOffice.OfficeApi.Enums.XlSeriesNameLevel SeriesNameLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SeriesNameLevel", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.XlSeriesNameLevel)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SeriesNameLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 15, 16)]
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
		/// SupportByVersion Office 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 15, 16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void UnProtect(object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password);
			Invoker.Method(this, "UnProtect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void UnProtect()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UnProtect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Protect()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Protect(object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Protect(object password, object drawingObjects)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Protect(object password, object drawingObjects, object contents)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object SeriesCollection(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "SeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object SeriesCollection()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "SeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void _ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey, autoText, hasLeaderLines);
			Invoker.Method(this, "_ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void _ApplyDataLabels()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void _ApplyDataLabels(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "_ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void _ApplyDataLabels(object type, object iMsoLegendKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey);
			Invoker.Method(this, "_ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void _ApplyDataLabels(object type, object iMsoLegendKey, object autoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey, autoText);
			Invoker.Method(this, "_ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		/// <param name="showPercentage">optional object ShowPercentage</param>
		/// <param name="showBubbleSize">optional object ShowBubbleSize</param>
		/// <param name="separator">optional object Separator</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyDataLabels()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyDataLabels(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyDataLabels(object type, object iMsoLegendKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyDataLabels(object type, object iMsoLegendKey, object autoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey, autoText);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey, autoText, hasLeaderLines);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		/// <param name="showPercentage">optional object ShowPercentage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="iMsoLegendKey">optional object IMsoLegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		/// <param name="showPercentage">optional object ShowPercentage</param>
		/// <param name="showBubbleSize">optional object ShowBubbleSize</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType ChartType</param>
		/// <param name="typeName">optional object TypeName</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chartType, typeName);
			Invoker.Method(this, "ApplyCustomType", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType ChartType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chartType);
			Invoker.Method(this, "ApplyCustomType", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="elementID">Int32 ElementID</param>
		/// <param name="arg1">Int32 Arg1</param>
		/// <param name="arg2">Int32 Arg2</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void GetChartElement(Int32 x, Int32 y, Int32 elementID, Int32 arg1, Int32 arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, elementID, arg1, arg2);
			Invoker.Method(this, "GetChartElement", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="source">string Source</param>
		/// <param name="plotBy">optional object PlotBy</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void SetSourceData(string source, object plotBy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, plotBy);
			Invoker.Method(this, "SetSourceData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="source">string Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void SetSourceData(string source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			Invoker.Method(this, "SetSourceData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional object Type</param>
		/// <param name="axisGroup">optional NetOffice.OfficeApi.Enums.XlAxisGroup AxisGroup = 1</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object Axes(object type, object axisGroup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, axisGroup);
			object returnItem = Invoker.MethodReturn(this, "Axes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object Axes()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Axes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object Axes(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "Axes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="rGallery">Int32 rGallery</param>
		/// <param name="varFormat">optional object varFormat</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void AutoFormat(Int32 rGallery, object varFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rGallery, varFormat);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="rGallery">Int32 rGallery</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void AutoFormat(Int32 rGallery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rGallery);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstr">string bstr</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void SetBackgroundPicture(string bstr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstr);
			Invoker.Method(this, "SetBackgroundPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varSource">optional object varSource</param>
		/// <param name="varGallery">optional object varGallery</param>
		/// <param name="varFormat">optional object varFormat</param>
		/// <param name="varPlotBy">optional object varPlotBy</param>
		/// <param name="varCategoryLabels">optional object varCategoryLabels</param>
		/// <param name="varSeriesLabels">optional object varSeriesLabels</param>
		/// <param name="varHasLegend">optional object varHasLegend</param>
		/// <param name="varTitle">optional object varTitle</param>
		/// <param name="varCategoryTitle">optional object varCategoryTitle</param>
		/// <param name="varValueTitle">optional object varValueTitle</param>
		/// <param name="varExtraTitle">optional object varExtraTitle</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle, object varValueTitle, object varExtraTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle, varValueTitle, varExtraTitle);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varSource">optional object varSource</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard(object varSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSource);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varSource">optional object varSource</param>
		/// <param name="varGallery">optional object varGallery</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard(object varSource, object varGallery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSource, varGallery);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varSource">optional object varSource</param>
		/// <param name="varGallery">optional object varGallery</param>
		/// <param name="varFormat">optional object varFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard(object varSource, object varGallery, object varFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSource, varGallery, varFormat);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varSource">optional object varSource</param>
		/// <param name="varGallery">optional object varGallery</param>
		/// <param name="varFormat">optional object varFormat</param>
		/// <param name="varPlotBy">optional object varPlotBy</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSource, varGallery, varFormat, varPlotBy);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varSource">optional object varSource</param>
		/// <param name="varGallery">optional object varGallery</param>
		/// <param name="varFormat">optional object varFormat</param>
		/// <param name="varPlotBy">optional object varPlotBy</param>
		/// <param name="varCategoryLabels">optional object varCategoryLabels</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varSource">optional object varSource</param>
		/// <param name="varGallery">optional object varGallery</param>
		/// <param name="varFormat">optional object varFormat</param>
		/// <param name="varPlotBy">optional object varPlotBy</param>
		/// <param name="varCategoryLabels">optional object varCategoryLabels</param>
		/// <param name="varSeriesLabels">optional object varSeriesLabels</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varSource">optional object varSource</param>
		/// <param name="varGallery">optional object varGallery</param>
		/// <param name="varFormat">optional object varFormat</param>
		/// <param name="varPlotBy">optional object varPlotBy</param>
		/// <param name="varCategoryLabels">optional object varCategoryLabels</param>
		/// <param name="varSeriesLabels">optional object varSeriesLabels</param>
		/// <param name="varHasLegend">optional object varHasLegend</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varSource">optional object varSource</param>
		/// <param name="varGallery">optional object varGallery</param>
		/// <param name="varFormat">optional object varFormat</param>
		/// <param name="varPlotBy">optional object varPlotBy</param>
		/// <param name="varCategoryLabels">optional object varCategoryLabels</param>
		/// <param name="varSeriesLabels">optional object varSeriesLabels</param>
		/// <param name="varHasLegend">optional object varHasLegend</param>
		/// <param name="varTitle">optional object varTitle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varSource">optional object varSource</param>
		/// <param name="varGallery">optional object varGallery</param>
		/// <param name="varFormat">optional object varFormat</param>
		/// <param name="varPlotBy">optional object varPlotBy</param>
		/// <param name="varCategoryLabels">optional object varCategoryLabels</param>
		/// <param name="varSeriesLabels">optional object varSeriesLabels</param>
		/// <param name="varHasLegend">optional object varHasLegend</param>
		/// <param name="varTitle">optional object varTitle</param>
		/// <param name="varCategoryTitle">optional object varCategoryTitle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varSource">optional object varSource</param>
		/// <param name="varGallery">optional object varGallery</param>
		/// <param name="varFormat">optional object varFormat</param>
		/// <param name="varPlotBy">optional object varPlotBy</param>
		/// <param name="varCategoryLabels">optional object varCategoryLabels</param>
		/// <param name="varSeriesLabels">optional object varSeriesLabels</param>
		/// <param name="varHasLegend">optional object varHasLegend</param>
		/// <param name="varTitle">optional object varTitle</param>
		/// <param name="varCategoryTitle">optional object varCategoryTitle</param>
		/// <param name="varValueTitle">optional object varValueTitle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle, object varValueTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle, varValueTitle);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="appearance">optional Int32 Appearance = 1</param>
		/// <param name="format">optional Int32 Format = -4147</param>
		/// <param name="size">optional Int32 Size = 2</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void CopyPicture(object appearance, object format, object size)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(appearance, format, size);
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void CopyPicture()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="appearance">optional Int32 Appearance = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void CopyPicture(object appearance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(appearance);
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="appearance">optional Int32 Appearance = 1</param>
		/// <param name="format">optional Int32 Format = -4147</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void CopyPicture(object appearance, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(appearance, format);
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varName">object varName</param>
		/// <param name="localeID">Int32 LocaleID</param>
		/// <param name="objType">Int32 ObjType</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object Evaluate(object varName, Int32 localeID, out Int32 objType)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			objType = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(varName, localeID, objType);
			object returnItem = Invoker.MethodReturn(this, "Evaluate", paramsArray, modifiers);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				objType = (Int32)paramsArray[2];
			return newObject;
			}
			else
			{
				objType = (Int32)paramsArray[2];
			return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varName">object varName</param>
		/// <param name="localeID">Int32 LocaleID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object _Evaluate(object varName, Int32 localeID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName, localeID);
			object returnItem = Invoker.MethodReturn(this, "_Evaluate", paramsArray);
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varType">optional object varType</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Paste(object varType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varType);
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstr">string bstr</param>
		/// <param name="varFilterName">optional object varFilterName</param>
		/// <param name="varInteractive">optional object varInteractive</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool Export(string bstr, object varFilterName, object varInteractive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstr, varFilterName, varInteractive);
			object returnItem = Invoker.MethodReturn(this, "Export", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstr">string bstr</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool Export(string bstr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstr);
			object returnItem = Invoker.MethodReturn(this, "Export", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstr">string bstr</param>
		/// <param name="varFilterName">optional object varFilterName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool Export(string bstr, object varFilterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstr, varFilterName);
			object returnItem = Invoker.MethodReturn(this, "Export", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="varName">object varName</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void SetDefaultChart(object varName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varName);
			Invoker.Method(this, "SetDefaultChart", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrFileName">string bstrFileName</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyChartTemplate(string bstrFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrFileName);
			Invoker.Method(this, "ApplyChartTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrFileName">string bstrFileName</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void SaveChartTemplate(string bstrFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrFileName);
			Invoker.Method(this, "SaveChartTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ClearToMatchStyle()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearToMatchStyle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void RefreshPivotTable()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RefreshPivotTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="layout">Int32 Layout</param>
		/// <param name="varChartType">optional object varChartType</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyLayout(Int32 layout, object varChartType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout, varChartType);
			Invoker.Method(this, "ApplyLayout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="layout">Int32 Layout</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ApplyLayout(Int32 layout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout);
			Invoker.Method(this, "ApplyLayout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Refresh()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Refresh", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="rHS">NetOffice.OfficeApi.Enums.MsoChartElementType RHS</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType rHS)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rHS);
			Invoker.Method(this, "SetElement", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object AreaGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "AreaGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object AreaGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AreaGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object BarGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "BarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object BarGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "BarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object ColumnGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "ColumnGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object ColumnGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ColumnGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object LineGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "LineGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object LineGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "LineGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object PieGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "PieGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object PieGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PieGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object DoughnutGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "DoughnutGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object DoughnutGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DoughnutGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object RadarGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "RadarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object RadarGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "RadarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object XYGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "XYGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object XYGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "XYGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public object Copy()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Copy", paramsArray);
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
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="replace">optional object Replace</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
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
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Office", 15, 16)]
		public object FullSeriesCollection(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "FullSeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 15, 16)]
		public object FullSeriesCollection()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FullSeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 15, 16)]
		public void DeleteHiddenContent()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DeleteHiddenContent", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 15, 16)]
		public void ClearToMatchColorStyle()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearToMatchColorStyle", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}