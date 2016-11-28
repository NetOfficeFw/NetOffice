using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.ExcelApi
{
	///<summary>
	/// _Chart
	///</summary>
	public class _Chart_ : COMObject
	{
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Chart_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart_() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// <param name="index2">optional object Index2</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840849.aspx
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object Index1</param>
        /// <param name="index2">optional object Index2</param>
        /// <param name="value">optional object value</param>
        [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_HasAxis(object index1, object index2, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index1, index2);
			Invoker.PropertySet(this, "HasAxis", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840849.aspx
		/// Alias for get_HasAxis
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// <param name="index2">optional object Index2</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object HasAxis(object index1, object index2)
		{
			return get_HasAxis(index1, index2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840849.aspx
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object Index1</param>
        /// <param name="value">optional object value</param>
        [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_HasAxis(object index1, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index1);
			Invoker.PropertySet(this, "HasAxis", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840849.aspx
		/// Alias for get_HasAxis
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object HasAxis(object index1)
		{
			return get_HasAxis(index1);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// DispatchInterface _Chart 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Chart : _Chart_
	{
		#pragma warning disable
		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
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
                    _type = typeof(_Chart);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Chart(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Chart(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838047.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.ExcelApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Application.LateBindingApiWrapperType) as NetOffice.ExcelApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195969.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCreator)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195815.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835278.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string CodeName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CodeName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string _CodeName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "_CodeName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "_CodeName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195753.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Index
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197207.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837108.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Next
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Next", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnDoubleClick
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OnDoubleClick", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OnDoubleClick", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnSheetActivate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OnSheetActivate", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OnSheetActivate", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnSheetDeactivate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OnSheetDeactivate", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OnSheetDeactivate", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836517.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PageSetup PageSetup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PageSetup", paramsArray);
				NetOffice.ExcelApi.PageSetup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.PageSetup.LateBindingApiWrapperType) as NetOffice.ExcelApi.PageSetup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838630.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Previous
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Previous", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193047.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ProtectContents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectContents", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822653.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ProtectDrawingObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectDrawingObjects", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821238.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ProtectionMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectionMode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839238.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSheetVisibility Visible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Visible", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlSheetVisibility)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Visible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823055.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shapes Shapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shapes", paramsArray);
				NetOffice.ExcelApi.Shapes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Shapes.LateBindingApiWrapperType) as NetOffice.ExcelApi.Shapes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ChartGroup Area3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Area3DGroup", paramsArray);
				NetOffice.ExcelApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.ExcelApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841256.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ChartGroup Bar3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Bar3DGroup", paramsArray);
				NetOffice.ExcelApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.ExcelApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194085.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ChartArea ChartArea
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartArea", paramsArray);
				NetOffice.ExcelApi.ChartArea newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ChartArea.LateBindingApiWrapperType) as NetOffice.ExcelApi.ChartArea;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196832.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ChartTitle ChartTitle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartTitle", paramsArray);
				NetOffice.ExcelApi.ChartTitle newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ChartTitle.LateBindingApiWrapperType) as NetOffice.ExcelApi.ChartTitle;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ChartGroup Column3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Column3DGroup", paramsArray);
				NetOffice.ExcelApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.ExcelApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Corners Corners
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Corners", paramsArray);
				NetOffice.ExcelApi.Corners newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Corners.LateBindingApiWrapperType) as NetOffice.ExcelApi.Corners;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840431.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.DataTable DataTable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataTable", paramsArray);
				NetOffice.ExcelApi.DataTable newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.DataTable.LateBindingApiWrapperType) as NetOffice.ExcelApi.DataTable;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196895.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838172.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlDisplayBlanksAs DisplayBlanksAs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayBlanksAs", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlDisplayBlanksAs)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayBlanksAs", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197517.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823205.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Floor Floor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Floor", paramsArray);
				NetOffice.ExcelApi.Floor newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Floor.LateBindingApiWrapperType) as NetOffice.ExcelApi.Floor;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821617.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840849.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838769.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840365.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836527.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837603.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198198.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Hyperlinks Hyperlinks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hyperlinks", paramsArray);
				NetOffice.ExcelApi.Hyperlinks newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Hyperlinks.LateBindingApiWrapperType) as NetOffice.ExcelApi.Hyperlinks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821884.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Legend Legend
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Legend", paramsArray);
				NetOffice.ExcelApi.Legend newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Legend.LateBindingApiWrapperType) as NetOffice.ExcelApi.Legend;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ChartGroup Line3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Line3DGroup", paramsArray);
				NetOffice.ExcelApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.ExcelApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196689.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ChartGroup Pie3DGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Pie3DGroup", paramsArray);
				NetOffice.ExcelApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.ExcelApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840927.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PlotArea PlotArea
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PlotArea", paramsArray);
				NetOffice.ExcelApi.PlotArea newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.PlotArea.LateBindingApiWrapperType) as NetOffice.ExcelApi.PlotArea;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840090.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821854.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838591.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool SizeWithWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SizeWithWindow", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SizeWithWindow", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ShowWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowWindow", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowWindow", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ChartGroup SurfaceGroup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SurfaceGroup", paramsArray);
				NetOffice.ExcelApi.ChartGroup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ChartGroup.LateBindingApiWrapperType) as NetOffice.ExcelApi.ChartGroup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820803.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlChartType ChartType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlChartType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ChartType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841192.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Walls Walls
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Walls", paramsArray);
				NetOffice.ExcelApi.Walls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Walls.LateBindingApiWrapperType) as NetOffice.ExcelApi.Walls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool WallsAndGridlines2D
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WallsAndGridlines2D", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WallsAndGridlines2D", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197600.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlBarShape BarShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BarShape", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlBarShape)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BarShape", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822363.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlRowCol PlotBy
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PlotBy", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlRowCol)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PlotBy", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822860.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195687.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837129.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838203.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotLayout PivotLayout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PivotLayout", paramsArray);
				NetOffice.ExcelApi.PivotLayout newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.PivotLayout.LateBindingApiWrapperType) as NetOffice.ExcelApi.PivotLayout;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Scripts Scripts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Scripts", paramsArray);
				NetOffice.OfficeApi.Scripts newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Scripts.LateBindingApiWrapperType) as NetOffice.OfficeApi.Scripts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838454.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Tab Tab
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tab", paramsArray);
				NetOffice.ExcelApi.Tab newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Tab.LateBindingApiWrapperType) as NetOffice.ExcelApi.Tab;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838210.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.MsoEnvelope MailEnvelope
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailEnvelope", paramsArray);
				NetOffice.OfficeApi.MsoEnvelope newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.MsoEnvelope.LateBindingApiWrapperType) as NetOffice.OfficeApi.MsoEnvelope;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194366.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834355.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Walls SideWall
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SideWall", paramsArray);
				NetOffice.ExcelApi.Walls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Walls.LateBindingApiWrapperType) as NetOffice.ExcelApi.Walls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838867.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Walls BackWall
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BackWall", paramsArray);
				NetOffice.ExcelApi.Walls newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Walls.LateBindingApiWrapperType) as NetOffice.ExcelApi.Walls;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838167.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835856.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 PrintedCommentPages
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintedCommentPages", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Dummy24
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dummy24", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Dummy24", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Dummy25
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dummy25", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Dummy25", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822505.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197522.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193279.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834352.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838192.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
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
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231310.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.Enums.XlCategoryLabelLevel CategoryLabelLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CategoryLabelLevel", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCategoryLabelLevel)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CategoryLabelLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227799.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.Enums.XlSeriesNameLevel SeriesNameLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SeriesNameLevel", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlSeriesNameLevel)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SeriesNameLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
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
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231021.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838025.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Activate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838866.aspx
		/// </summary>
		/// <param name="before">optional object Before</param>
		/// <param name="after">optional object After</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Copy(object before, object after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(before, after);
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838866.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838866.aspx
		/// </summary>
		/// <param name="before">optional object Before</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Copy(object before)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(before);
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822797.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840583.aspx
		/// </summary>
		/// <param name="before">optional object Before</param>
		/// <param name="after">optional object After</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Move(object before, object after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(before, after);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840583.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Move()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840583.aspx
		/// </summary>
		/// <param name="before">optional object Before</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Move(object before)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(before);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile, collate);
			Invoker.Method(this, "_PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="prToFileName">optional object PrToFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName);
			Invoker.Method(this, "_PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from);
			Invoker.Method(this, "_PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to);
			Invoker.Method(this, "_PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies);
			Invoker.Method(this, "_PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview);
			Invoker.Method(this, "_PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter);
			Invoker.Method(this, "_PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile);
			Invoker.Method(this, "_PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838625.aspx
		/// </summary>
		/// <param name="enableChanges">optional object EnableChanges</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PrintPreview(object enableChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(enableChanges);
			Invoker.Method(this, "PrintPreview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838625.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PrintPreview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintPreview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821561.aspx
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821561.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Protect()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821561.aspx
		/// </summary>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821561.aspx
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object drawingObjects)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821561.aspx
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object drawingObjects, object contents)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821561.aspx
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy23()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_Dummy23", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195059.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="textCodepage">optional object TextCodepage</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195059.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="textCodepage">optional object TextCodepage</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		/// <param name="local">optional object Local</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout, object local)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout, local);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195059.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195059.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195059.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195059.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195059.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195059.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195059.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="addToMru">optional object AddToMru</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195059.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="textCodepage">optional object TextCodepage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195030.aspx
		/// </summary>
		/// <param name="replace">optional object Replace</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Select(object replace)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(replace);
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195030.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Select()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821208.aspx
		/// </summary>
		/// <param name="password">optional object Password</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Unprotect(object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password);
			Invoker.Method(this, "Unprotect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821208.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Unprotect()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Unprotect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195014.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195014.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		/// <param name="showPercentage">optional object ShowPercentage</param>
		/// <param name="showBubbleSize">optional object ShowBubbleSize</param>
		/// <param name="separator">optional object Separator</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195014.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ApplyDataLabels()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195014.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ApplyDataLabels(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195014.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ApplyDataLabels(object type, object legendKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195014.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195014.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195014.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195014.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195014.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		/// <param name="showPercentage">optional object ShowPercentage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195014.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		/// <param name="showSeriesName">optional object ShowSeriesName</param>
		/// <param name="showCategoryName">optional object ShowCategoryName</param>
		/// <param name="showValue">optional object ShowValue</param>
		/// <param name="showPercentage">optional object ShowPercentage</param>
		/// <param name="showBubbleSize">optional object ShowBubbleSize</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize);
			Invoker.Method(this, "ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Arcs(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Arcs", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Arcs()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Arcs", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AreaGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "AreaGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AreaGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AreaGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="gallery">Int32 Gallery</param>
		/// <param name="format">optional object Format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void AutoFormat(Int32 gallery, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gallery, format);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="gallery">Int32 Gallery</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void AutoFormat(Int32 gallery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gallery);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839276.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		/// <param name="axisGroup">optional NetOffice.ExcelApi.Enums.XlAxisGroup AxisGroup = 1</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Axes(object type, object axisGroup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, axisGroup);
			object returnItem = Invoker.MethodReturn(this, "Axes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839276.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Axes()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Axes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839276.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Axes(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "Axes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194060.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SetBackgroundPicture(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "SetBackgroundPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object BarGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "BarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object BarGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "BarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Buttons(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Buttons", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Buttons()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Buttons", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840069.aspx
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ChartGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "ChartGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840069.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ChartGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ChartGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821276.aspx
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ChartObjects(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "ChartObjects", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821276.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ChartObjects()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ChartObjects", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
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
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle, object extraTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle, extraTitle);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard(object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard(object source, object gallery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard(object source, object gallery, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		/// <param name="seriesLabels">optional object SeriesLabels</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="gallery">optional object Gallery</param>
		/// <param name="format">optional object Format</param>
		/// <param name="plotBy">optional object PlotBy</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		/// <param name="seriesLabels">optional object SeriesLabels</param>
		/// <param name="hasLegend">optional object HasLegend</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
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
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
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
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838804.aspx
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
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle);
			Invoker.Method(this, "ChartWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CheckBoxes(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "CheckBoxes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CheckBoxes()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CheckBoxes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836772.aspx
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="spellLang">optional object SpellLang</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, spellLang);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836772.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CheckSpelling()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836772.aspx
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836772.aspx
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836772.aspx
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest);
			Invoker.Method(this, "CheckSpelling", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ColumnGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "ColumnGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ColumnGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ColumnGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841052.aspx
		/// </summary>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlCopyPictureFormat Format = -4147</param>
		/// <param name="size">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Size = 2</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CopyPicture(object appearance, object format, object size)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(appearance, format, size);
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841052.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CopyPicture()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841052.aspx
		/// </summary>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CopyPicture(object appearance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(appearance);
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841052.aspx
		/// </summary>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlCopyPictureFormat Format = -4147</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CopyPicture(object appearance, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(appearance, format);
			Invoker.Method(this, "CopyPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="size">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Size = 1</param>
		/// <param name="containsPICT">optional object ContainsPICT</param>
		/// <param name="containsBIFF">optional object ContainsBIFF</param>
		/// <param name="containsRTF">optional object ContainsRTF</param>
		/// <param name="containsVALU">optional object ContainsVALU</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CreatePublisher(object edition, object appearance, object size, object containsPICT, object containsBIFF, object containsRTF, object containsVALU)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, appearance, size, containsPICT, containsBIFF, containsRTF, containsVALU);
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CreatePublisher()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CreatePublisher(object edition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition);
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CreatePublisher(object edition, object appearance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, appearance);
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="size">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Size = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CreatePublisher(object edition, object appearance, object size)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, appearance, size);
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="size">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Size = 1</param>
		/// <param name="containsPICT">optional object ContainsPICT</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CreatePublisher(object edition, object appearance, object size, object containsPICT)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, appearance, size, containsPICT);
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="size">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Size = 1</param>
		/// <param name="containsPICT">optional object ContainsPICT</param>
		/// <param name="containsBIFF">optional object ContainsBIFF</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CreatePublisher(object edition, object appearance, object size, object containsPICT, object containsBIFF)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, appearance, size, containsPICT, containsBIFF);
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="size">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Size = 1</param>
		/// <param name="containsPICT">optional object ContainsPICT</param>
		/// <param name="containsBIFF">optional object ContainsBIFF</param>
		/// <param name="containsRTF">optional object ContainsRTF</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CreatePublisher(object edition, object appearance, object size, object containsPICT, object containsBIFF, object containsRTF)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, appearance, size, containsPICT, containsBIFF, containsRTF);
			Invoker.Method(this, "CreatePublisher", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Deselect()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Deselect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DoughnutGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "DoughnutGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DoughnutGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DoughnutGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Drawings(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Drawings", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Drawings()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Drawings", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DrawingObjects(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "DrawingObjects", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DrawingObjects()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DrawingObjects", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DropDowns(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "DropDowns", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DropDowns()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DropDowns", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834376.aspx
		/// </summary>
		/// <param name="name">object Name</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Evaluate(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "Evaluate", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">object Name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Evaluate(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GroupBoxes(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "GroupBoxes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GroupBoxes()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GroupBoxes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GroupObjects(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "GroupObjects", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GroupObjects()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GroupObjects", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Labels(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Labels", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Labels()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Labels", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object LineGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "LineGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object LineGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "LineGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Lines(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Lines", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Lines()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Lines", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ListBoxes(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "ListBoxes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ListBoxes()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ListBoxes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196573.aspx
		/// </summary>
		/// <param name="where">NetOffice.ExcelApi.Enums.XlChartLocation Where</param>
		/// <param name="name">optional object Name</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Chart Location(NetOffice.ExcelApi.Enums.XlChartLocation where, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(where, name);
			object returnItem = Invoker.MethodReturn(this, "Location", paramsArray);
			NetOffice.ExcelApi.Chart newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Chart.LateBindingApiWrapperType) as NetOffice.ExcelApi.Chart;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196573.aspx
		/// </summary>
		/// <param name="where">NetOffice.ExcelApi.Enums.XlChartLocation Where</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Chart Location(NetOffice.ExcelApi.Enums.XlChartLocation where)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(where);
			object returnItem = Invoker.MethodReturn(this, "Location", paramsArray);
			NetOffice.ExcelApi.Chart newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Chart.LateBindingApiWrapperType) as NetOffice.ExcelApi.Chart;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840253.aspx
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object OLEObjects(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "OLEObjects", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840253.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object OLEObjects()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "OLEObjects", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object OptionButtons(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "OptionButtons", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object OptionButtons()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "OptionButtons", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Ovals(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Ovals", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Ovals()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Ovals", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840204.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Paste(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840204.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Pictures(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Pictures", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Pictures()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Pictures", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PieGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "PieGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PieGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PieGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object RadarGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "RadarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object RadarGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "RadarGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Rectangles(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Rectangles", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Rectangles()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Rectangles", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ScrollBars(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "ScrollBars", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ScrollBars()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ScrollBars", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193558.aspx
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SeriesCollection(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "SeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193558.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SeriesCollection()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "SeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Spinners(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Spinners", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Spinners()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Spinners", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextBoxes(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "TextBoxes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextBoxes()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "TextBoxes", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="chartType">NetOffice.ExcelApi.Enums.XlChartType ChartType</param>
		/// <param name="typeName">optional object TypeName</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ApplyCustomType(NetOffice.ExcelApi.Enums.XlChartType chartType, object typeName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chartType, typeName);
			Invoker.Method(this, "ApplyCustomType", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="chartType">NetOffice.ExcelApi.Enums.XlChartType ChartType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ApplyCustomType(NetOffice.ExcelApi.Enums.XlChartType chartType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chartType);
			Invoker.Method(this, "ApplyCustomType", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object XYGroups(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "XYGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object XYGroups()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "XYGroups", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CopyChartBuild()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CopyChartBuild", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837393.aspx
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="elementID">Int32 ElementID</param>
		/// <param name="arg1">Int32 Arg1</param>
		/// <param name="arg2">Int32 Arg2</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void GetChartElement(Int32 x, Int32 y, Int32 elementID, Int32 arg1, Int32 arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, elementID, arg1, arg2);
			Invoker.Method(this, "GetChartElement", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841196.aspx
		/// </summary>
		/// <param name="source">NetOffice.ExcelApi.Range Source</param>
		/// <param name="plotBy">optional object PlotBy</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SetSourceData(NetOffice.ExcelApi.Range source, object plotBy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, plotBy);
			Invoker.Method(this, "SetSourceData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841196.aspx
		/// </summary>
		/// <param name="source">NetOffice.ExcelApi.Range Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SetSourceData(NetOffice.ExcelApi.Range source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			Invoker.Method(this, "SetSourceData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198129.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="interactive">optional object Interactive</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Export(string filename, object filterName, object interactive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, filterName, interactive);
			object returnItem = Invoker.MethodReturn(this, "Export", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198129.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Export(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			object returnItem = Invoker.MethodReturn(this, "Export", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198129.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="filterName">optional object FilterName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Export(string filename, object filterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, filterName);
			object returnItem = Invoker.MethodReturn(this, "Export", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198180.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Refresh()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Refresh", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821945.aspx
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="prToFileName">optional object PrToFileName</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821945.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821945.aspx
		/// </summary>
		/// <param name="from">optional object From</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821945.aspx
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821945.aspx
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821945.aspx
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821945.aspx
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821945.aspx
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821945.aspx
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile, collate);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly);
			Invoker.Method(this, "_Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _Protect()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _Protect(object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password);
			Invoker.Method(this, "_Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _Protect(object password, object drawingObjects)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects);
			Invoker.Method(this, "_Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _Protect(object password, object drawingObjects, object contents)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents);
			Invoker.Method(this, "_Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _Protect(object password, object drawingObjects, object contents, object scenarios)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios);
			Invoker.Method(this, "_Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="textCodepage">optional object TextCodepage</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(string filename, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(string filename, object fileFormat, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="addToMru">optional object AddToMru</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="textCodepage">optional object TextCodepage</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		/// <param name="hasLeaderLines">optional object HasLeaderLines</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText, hasLeaderLines);
			Invoker.Method(this, "_ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _ApplyDataLabels()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _ApplyDataLabels(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "_ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _ApplyDataLabels(object type, object legendKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey);
			Invoker.Method(this, "_ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object LegendKey</param>
		/// <param name="autoText">optional object AutoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _ApplyDataLabels(object type, object legendKey, object autoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, legendKey, autoText);
			Invoker.Method(this, "_ApplyDataLabels", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile, collate);
			Invoker.Method(this, "__PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void __PrintOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "__PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void __PrintOut(object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from);
			Invoker.Method(this, "__PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to);
			Invoker.Method(this, "__PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies);
			Invoker.Method(this, "__PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview);
			Invoker.Method(this, "__PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter);
			Invoker.Method(this, "__PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile);
			Invoker.Method(this, "__PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193792.aspx
		/// </summary>
		/// <param name="layout">Int32 Layout</param>
		/// <param name="chartType">optional object ChartType</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ApplyLayout(Int32 layout, object chartType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout, chartType);
			Invoker.Method(this, "ApplyLayout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193792.aspx
		/// </summary>
		/// <param name="layout">Int32 Layout</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ApplyLayout(Int32 layout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout);
			Invoker.Method(this, "ApplyLayout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193844.aspx
		/// </summary>
		/// <param name="element">NetOffice.OfficeApi.Enums.MsoChartElementType Element</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType element)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(element);
			Invoker.Method(this, "SetElement", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838076.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ApplyChartTemplate(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "ApplyChartTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839779.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void SaveChartTemplate(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "SaveChartTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835564.aspx
		/// </summary>
		/// <param name="name">object Name</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void SetDefaultChart(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "SetDefaultChart", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198218.aspx
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object IgnorePrintAreas</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="openAfterPublish">optional object OpenAfterPublish</param>
		/// <param name="fixedFormatExtClassPtr">optional object FixedFormatExtClassPtr</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish, object fixedFormatExtClassPtr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, fixedFormatExtClassPtr);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198218.aspx
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198218.aspx
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198218.aspx
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198218.aspx
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198218.aspx
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object IgnorePrintAreas</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties, ignorePrintAreas);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198218.aspx
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object IgnorePrintAreas</param>
		/// <param name="from">optional object From</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties, ignorePrintAreas, from);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198218.aspx
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object IgnorePrintAreas</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198218.aspx
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object IgnorePrintAreas</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="openAfterPublish">optional object OpenAfterPublish</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835627.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ClearToMatchStyle()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearToMatchStyle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230578.aspx
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public object FullSeriesCollection(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "FullSeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230578.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 15, 16)]
		public object FullSeriesCollection()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FullSeriesCollection", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 15, 16)]
		public void DeleteHiddenContent()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DeleteHiddenContent", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229445.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public void ClearToMatchColorStyle()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearToMatchColorStyle", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}