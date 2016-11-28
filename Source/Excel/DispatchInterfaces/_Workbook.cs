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
	/// _Workbook
	///</summary>
	public class _Workbook_ : COMObject
	{
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Workbook_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Colors(object index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Colors", paramsArray);
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
        /// <param name="index">optional object Index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Colors(object index, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.PropertySet(this, "Colors", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx
		/// Alias for get_Colors
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Colors(object index)
		{
			return get_Colors(index);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// DispatchInterface _Workbook 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Workbook : _Workbook_
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
                    _type = typeof(_Workbook);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Workbook(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835918.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840080.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198008.aspx
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
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool AcceptLabelsInFormulas
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AcceptLabelsInFormulas", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AcceptLabelsInFormulas", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834923.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Chart ActiveChart
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveChart", paramsArray);
				NetOffice.ExcelApi.Chart newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Chart.LateBindingApiWrapperType) as NetOffice.ExcelApi.Chart;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841181.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ActiveSheet
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveSheet", paramsArray);
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
		public string Author
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Author", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Author", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840067.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 AutoUpdateFrequency
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoUpdateFrequency", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutoUpdateFrequency", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193298.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool AutoUpdateSaveChanges
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoUpdateSaveChanges", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutoUpdateSaveChanges", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821530.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 ChangeHistoryDuration
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChangeHistoryDuration", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ChangeHistoryDuration", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197172.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object BuiltinDocumentProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BuiltinDocumentProperties", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821062.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Sheets Charts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Charts", paramsArray);
				NetOffice.ExcelApi.Sheets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType) as NetOffice.ExcelApi.Sheets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195162.aspx
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
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Colors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Colors", paramsArray);
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
				Invoker.PropertySet(this, "Colors", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835614.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandBars", paramsArray);
				NetOffice.OfficeApi.CommandBars newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Comments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Comments", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Comments", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198339.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSaveConflictResolution ConflictResolution
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConflictResolution", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlSaveConflictResolution)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ConflictResolution", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834401.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Container
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Container", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196337.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool CreateBackup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CreateBackup", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834990.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CustomDocumentProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomDocumentProperties", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193264.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Date1904
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Date1904", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Date1904", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Sheets DialogSheets
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DialogSheets", paramsArray);
				NetOffice.ExcelApi.Sheets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType) as NetOffice.ExcelApi.Sheets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834329.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.xlDisplayDrawingObjects DisplayDrawingObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayDrawingObjects", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.xlDisplayDrawingObjects)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayDrawingObjects", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840717.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlFileFormat FileFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileFormat", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlFileFormat)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834975.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string FullName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FullName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool HasMailer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasMailer", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HasMailer", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840238.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool HasPassword
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasPassword", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool HasRoutingSlip
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasRoutingSlip", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HasRoutingSlip", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838249.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool IsAddin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsAddin", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IsAddin", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Keywords
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Keywords", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Keywords", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837965.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Mailer Mailer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Mailer", paramsArray);
				NetOffice.ExcelApi.Mailer newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Mailer.LateBindingApiWrapperType) as NetOffice.ExcelApi.Mailer;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Sheets Modules
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Modules", paramsArray);
				NetOffice.ExcelApi.Sheets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType) as NetOffice.ExcelApi.Sheets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839882.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool MultiUserEditing
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MultiUserEditing", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820899.aspx
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
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195422.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Names Names
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Names", paramsArray);
				NetOffice.ExcelApi.Names newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Names.LateBindingApiWrapperType) as NetOffice.ExcelApi.Names;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnSave
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OnSave", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OnSave", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840974.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string Path
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Path", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836500.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool PersonalViewListSettings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PersonalViewListSettings", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PersonalViewListSettings", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822649.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool PersonalViewPrintSettings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PersonalViewPrintSettings", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PersonalViewPrintSettings", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198189.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool PrecisionAsDisplayed
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrecisionAsDisplayed", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrecisionAsDisplayed", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838601.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ProtectStructure
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectStructure", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193864.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ProtectWindows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectWindows", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840925.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ReadOnly
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadOnly", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196964.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ReadOnlyRecommended
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadOnlyRecommended", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ReadOnlyRecommended", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834665.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 RevisionNumber
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RevisionNumber", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Routed
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Routed", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.RoutingSlip RoutingSlip
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RoutingSlip", paramsArray);
				NetOffice.ExcelApi.RoutingSlip newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.RoutingSlip.LateBindingApiWrapperType) as NetOffice.ExcelApi.RoutingSlip;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196613.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Saved
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Saved", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Saved", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840667.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool SaveLinkValues
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SaveLinkValues", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SaveLinkValues", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197568.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Sheets Sheets
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sheets", paramsArray);
				NetOffice.ExcelApi.Sheets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType) as NetOffice.ExcelApi.Sheets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839677.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ShowConflictHistory
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowConflictHistory", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowConflictHistory", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839039.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Styles Styles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Styles", paramsArray);
				NetOffice.ExcelApi.Styles newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Styles.LateBindingApiWrapperType) as NetOffice.ExcelApi.Styles;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Subject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Subject", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Subject", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Title
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Title", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Title", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193266.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool UpdateRemoteReferences
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UpdateRemoteReferences", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UpdateRemoteReferences", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool UserControl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UserControl", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UserControl", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193788.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object UserStatus
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UserStatus", paramsArray);
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
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195531.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CustomViews CustomViews
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomViews", paramsArray);
				NetOffice.ExcelApi.CustomViews newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.CustomViews.LateBindingApiWrapperType) as NetOffice.ExcelApi.CustomViews;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195152.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Windows Windows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Windows", paramsArray);
				NetOffice.ExcelApi.Windows newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Windows.LateBindingApiWrapperType) as NetOffice.ExcelApi.Windows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835542.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Sheets Worksheets
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Worksheets", paramsArray);
				NetOffice.ExcelApi.Sheets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType) as NetOffice.ExcelApi.Sheets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836228.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool WriteReserved
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WriteReserved", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840737.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string WriteReservedBy
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WriteReservedBy", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822819.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Sheets Excel4IntlMacroSheets
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Excel4IntlMacroSheets", paramsArray);
				NetOffice.ExcelApi.Sheets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType) as NetOffice.ExcelApi.Sheets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195645.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Sheets Excel4MacroSheets
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Excel4MacroSheets", paramsArray);
				NetOffice.ExcelApi.Sheets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType) as NetOffice.ExcelApi.Sheets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836472.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool TemplateRemoveExtData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TemplateRemoveExtData", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TemplateRemoveExtData", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194254.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool HighlightChangesOnScreen
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HighlightChangesOnScreen", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HighlightChangesOnScreen", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197016.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool KeepChangeHistory
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "KeepChangeHistory", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "KeepChangeHistory", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834301.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ListChangesOnNewSheet
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListChangesOnNewSheet", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ListChangesOnNewSheet", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194737.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBProject VBProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBProject", paramsArray);
				NetOffice.VBIDEApi.VBProject newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.VBProject.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840963.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool IsInplace
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsInplace", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838208.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PublishObjects PublishObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PublishObjects", paramsArray);
				NetOffice.ExcelApi.PublishObjects newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.PublishObjects.LateBindingApiWrapperType) as NetOffice.ExcelApi.PublishObjects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834724.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.WebOptions WebOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WebOptions", paramsArray);
				NetOffice.ExcelApi.WebOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.WebOptions.LateBindingApiWrapperType) as NetOffice.ExcelApi.WebOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.HTMLProject HTMLProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HTMLProject", paramsArray);
				NetOffice.OfficeApi.HTMLProject newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.HTMLProject.LateBindingApiWrapperType) as NetOffice.OfficeApi.HTMLProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839554.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool EnvelopeVisible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnvelopeVisible", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnvelopeVisible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193512.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 CalculationVersion
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CalculationVersion", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822659.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool VBASigned
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBASigned", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool _ReadOnlyRecommended
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "_ReadOnlyRecommended", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196322.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool ShowPivotTableFieldList
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowPivotTableFieldList", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowPivotTableFieldList", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839021.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlUpdateLinks UpdateLinks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UpdateLinks", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlUpdateLinks)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UpdateLinks", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193225.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool EnableAutoRecover
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableAutoRecover", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableAutoRecover", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841017.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool RemovePersonalInformation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RemovePersonalInformation", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RemovePersonalInformation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821089.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public string FullNameURLEncoded
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FullNameURLEncoded", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821529.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public string Password
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Password", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Password", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837767.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public string WritePassword
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WritePassword", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WritePassword", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839579.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public string PasswordEncryptionProvider
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionProvider", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195464.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public string PasswordEncryptionAlgorithm
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionAlgorithm", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195381.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 PasswordEncryptionKeyLength
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionKeyLength", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820819.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool PasswordEncryptionFileProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionFileProperties", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.SmartTagOptions SmartTagOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartTagOptions", paramsArray);
				NetOffice.ExcelApi.SmartTagOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.SmartTagOptions.LateBindingApiWrapperType) as NetOffice.ExcelApi.SmartTagOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840697.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Permission Permission
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Permission", paramsArray);
				NetOffice.OfficeApi.Permission newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Permission.LateBindingApiWrapperType) as NetOffice.OfficeApi.Permission;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835236.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SharedWorkspace", paramsArray);
				NetOffice.OfficeApi.SharedWorkspace newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SharedWorkspace.LateBindingApiWrapperType) as NetOffice.OfficeApi.SharedWorkspace;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192923.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Sync Sync
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sync", paramsArray);
				NetOffice.OfficeApi.Sync newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Sync.LateBindingApiWrapperType) as NetOffice.OfficeApi.Sync;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838260.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.XmlNamespaces XmlNamespaces
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XmlNamespaces", paramsArray);
				NetOffice.ExcelApi.XmlNamespaces newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.XmlNamespaces.LateBindingApiWrapperType) as NetOffice.ExcelApi.XmlNamespaces;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838975.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.XmlMaps XmlMaps
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XmlMaps", paramsArray);
				NetOffice.ExcelApi.XmlMaps newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.XmlMaps.LateBindingApiWrapperType) as NetOffice.ExcelApi.XmlMaps;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194561.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SmartDocument SmartDocument
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartDocument", paramsArray);
				NetOffice.OfficeApi.SmartDocument newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartDocument.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartDocument;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838205.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentLibraryVersions", paramsArray);
				NetOffice.OfficeApi.DocumentLibraryVersions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.DocumentLibraryVersions.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentLibraryVersions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837429.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public bool InactiveListBorderVisible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InactiveListBorderVisible", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "InactiveListBorderVisible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838435.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public bool DisplayInkComments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayInkComments", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayInkComments", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837152.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.MetaProperties ContentTypeProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContentTypeProperties", paramsArray);
				NetOffice.OfficeApi.MetaProperties newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.MetaProperties.LateBindingApiWrapperType) as NetOffice.OfficeApi.MetaProperties;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836773.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Connections Connections
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Connections", paramsArray);
				NetOffice.ExcelApi.Connections newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Connections.LateBindingApiWrapperType) as NetOffice.ExcelApi.Connections;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838073.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.SignatureSet Signatures
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Signatures", paramsArray);
				NetOffice.OfficeApi.SignatureSet newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SignatureSet.LateBindingApiWrapperType) as NetOffice.OfficeApi.SignatureSet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194489.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.ServerPolicy ServerPolicy
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ServerPolicy", paramsArray);
				NetOffice.OfficeApi.ServerPolicy newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.ServerPolicy.LateBindingApiWrapperType) as NetOffice.OfficeApi.ServerPolicy;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195426.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.DocumentInspectors DocumentInspectors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentInspectors", paramsArray);
				NetOffice.OfficeApi.DocumentInspectors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.DocumentInspectors.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentInspectors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195818.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ServerViewableItems ServerViewableItems
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ServerViewableItems", paramsArray);
				NetOffice.ExcelApi.ServerViewableItems newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ServerViewableItems.LateBindingApiWrapperType) as NetOffice.ExcelApi.ServerViewableItems;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837756.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.TableStyles TableStyles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TableStyles", paramsArray);
				NetOffice.ExcelApi.TableStyles newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.TableStyles.LateBindingApiWrapperType) as NetOffice.ExcelApi.TableStyles;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195934.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DefaultTableStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultTableStyle", paramsArray);
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
				Invoker.PropertySet(this, "DefaultTableStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835624.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DefaultPivotTableStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultPivotTableStyle", paramsArray);
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
				Invoker.PropertySet(this, "DefaultPivotTableStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836165.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool CheckCompatibility
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CheckCompatibility", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CheckCompatibility", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838063.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool HasVBProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasVBProject", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838448.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLParts CustomXMLParts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomXMLParts", paramsArray);
				NetOffice.OfficeApi.CustomXMLParts newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLParts.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLParts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820907.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool Final
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Final", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Final", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196847.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Research Research
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Research", paramsArray);
				NetOffice.ExcelApi.Research newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Research.LateBindingApiWrapperType) as NetOffice.ExcelApi.Research;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194072.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.OfficeTheme Theme
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Theme", paramsArray);
				NetOffice.OfficeApi.OfficeTheme newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.OfficeTheme.LateBindingApiWrapperType) as NetOffice.OfficeApi.OfficeTheme;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834991.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool Excel8CompatibilityMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Excel8CompatibilityMode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837960.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool ConnectionsDisabled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConnectionsDisabled", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835280.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool ShowPivotChartActiveFields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowPivotChartActiveFields", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowPivotChartActiveFields", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839003.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.IconSets IconSets
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IconSets", paramsArray);
				NetOffice.ExcelApi.IconSets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.IconSets.LateBindingApiWrapperType) as NetOffice.ExcelApi.IconSets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194147.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public string EncryptionProvider
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EncryptionProvider", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EncryptionProvider", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839440.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool DoNotPromptForConvert
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DoNotPromptForConvert", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DoNotPromptForConvert", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823189.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool ForceFullCalculation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ForceFullCalculation", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ForceFullCalculation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194925.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SlicerCaches SlicerCaches
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SlicerCaches", paramsArray);
				NetOffice.ExcelApi.SlicerCaches newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.SlicerCaches.LateBindingApiWrapperType) as NetOffice.ExcelApi.SlicerCaches;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839464.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer ActiveSlicer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveSlicer", paramsArray);
				NetOffice.ExcelApi.Slicer newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType) as NetOffice.ExcelApi.Slicer;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193862.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DefaultSlicerStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultSlicerStyle", paramsArray);
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
				Invoker.PropertySet(this, "DefaultSlicerStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838425.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 AccuracyVersion
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AccuracyVersion", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AccuracyVersion", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229542.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public bool CaseSensitive
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CaseSensitive", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231362.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public bool UseWholeCellCriteria
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UseWholeCellCriteria", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230772.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public bool UseWildcards
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UseWildcards", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231277.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public object PivotTables
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PivotTables", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228926.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.Model Model
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Model", paramsArray);
				NetOffice.ExcelApi.Model newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Model.LateBindingApiWrapperType) as NetOffice.ExcelApi.Model;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227452.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public bool ChartDataPointTrack
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartDataPointTrack", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ChartDataPointTrack", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230214.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DefaultTimelineStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultTimelineStyle", paramsArray);
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
				Invoker.PropertySet(this, "DefaultTimelineStyle", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821837.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Activate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx
		/// </summary>
		/// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess Mode</param>
		/// <param name="writePassword">optional object WritePassword</param>
		/// <param name="notify">optional object Notify</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode, object writePassword, object notify)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(mode, writePassword, notify);
			Invoker.Method(this, "ChangeFileAccess", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx
		/// </summary>
		/// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess Mode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(mode);
			Invoker.Method(this, "ChangeFileAccess", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx
		/// </summary>
		/// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess Mode</param>
		/// <param name="writePassword">optional object WritePassword</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode, object writePassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(mode, writePassword);
			Invoker.Method(this, "ChangeFileAccess", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836537.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="newName">string NewName</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlLinkType Type = 1</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChangeLink(string name, string newName, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, newName, type);
			Invoker.Method(this, "ChangeLink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836537.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="newName">string NewName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ChangeLink(string name, string newName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, newName);
			Invoker.Method(this, "ChangeLink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="routeWorkbook">optional object RouteWorkbook</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges, object filename, object routeWorkbook)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, filename, routeWorkbook);
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="filename">optional object Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges, object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, filename);
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839565.aspx
		/// </summary>
		/// <param name="numberFormat">string NumberFormat</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void DeleteNumberFormat(string numberFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numberFormat);
			Invoker.Method(this, "DeleteNumberFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836762.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ExclusiveAccess()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ExclusiveAccess", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836208.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ForwardMailer()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ForwardMailer", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo LinkInfo</param>
		/// <param name="type">optional object Type</param>
		/// <param name="editionRef">optional object EditionRef</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo, object type, object editionRef)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, linkInfo, type, editionRef);
			object returnItem = Invoker.MethodReturn(this, "LinkInfo", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo LinkInfo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, linkInfo);
			object returnItem = Invoker.MethodReturn(this, "LinkInfo", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo LinkInfo</param>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, linkInfo, type);
			object returnItem = Invoker.MethodReturn(this, "LinkInfo", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821922.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object LinkSources(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "LinkSources", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821922.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object LinkSources()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "LinkSources", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196693.aspx
		/// </summary>
		/// <param name="filename">object Filename</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MergeWorkbook(object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "MergeWorkbook", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838378.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Window NewWindow()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NewWindow", paramsArray);
			NetOffice.ExcelApi.Window newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Window.LateBindingApiWrapperType) as NetOffice.ExcelApi.Window;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="type">optional object Type</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenLinks(string name, object readOnly, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, readOnly, type);
			Invoker.Method(this, "OpenLinks", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenLinks(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "OpenLinks", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenLinks(string name, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, readOnly);
			Invoker.Method(this, "OpenLinks", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193549.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotCaches PivotCaches()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PivotCaches", paramsArray);
			NetOffice.ExcelApi.PivotCaches newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.PivotCaches.LateBindingApiWrapperType) as NetOffice.ExcelApi.PivotCaches;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821844.aspx
		/// </summary>
		/// <param name="destName">optional object DestName</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Post(object destName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destName);
			Invoker.Method(this, "Post", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821844.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Post()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Post", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193068.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193068.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="structure">optional object Structure</param>
		/// <param name="windows">optional object Windows</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object structure, object windows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, structure, windows);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="structure">optional object Structure</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object structure)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, structure);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="sharingPassword">optional object SharingPassword</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, password, writeResPassword, readOnlyRecommended, createBackup, sharingPassword);
			Invoker.Method(this, "ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="sharingPassword">optional object SharingPassword</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, password, writeResPassword, readOnlyRecommended, createBackup, sharingPassword, fileFormat);
			Invoker.Method(this, "ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, password);
			Invoker.Method(this, "ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename, object password, object writeResPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, password, writeResPassword);
			Invoker.Method(this, "ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, password, writeResPassword, readOnlyRecommended);
			Invoker.Method(this, "ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, password, writeResPassword, readOnlyRecommended, createBackup);
			Invoker.Method(this, "ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838648.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void RefreshAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RefreshAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820902.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Reply()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Reply", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838788.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ReplyAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReplyAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840747.aspx
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void RemoveUser(Int32 index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "RemoveUser", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Route()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Route", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835203.aspx
		/// </summary>
		/// <param name="which">NetOffice.ExcelApi.Enums.XlRunAutoMacro Which</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void RunAutoMacros(NetOffice.ExcelApi.Enums.XlRunAutoMacro which)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(which);
			Invoker.Method(this, "RunAutoMacros", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197585.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Save()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object ConflictResolution</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="textCodepage">optional object TextCodepage</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object ConflictResolution</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="textCodepage">optional object TextCodepage</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		/// <param name="local">optional object Local</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout, object local)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout, local);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object ConflictResolution</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object ConflictResolution</param>
		/// <param name="addToMru">optional object AddToMru</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object ConflictResolution</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="textCodepage">optional object TextCodepage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835014.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveCopyAs(object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835014.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveCopyAs()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx
		/// </summary>
		/// <param name="recipients">object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		/// <param name="returnReceipt">optional object ReturnReceipt</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SendMail(object recipients, object subject, object returnReceipt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, returnReceipt);
			Invoker.Method(this, "SendMail", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx
		/// </summary>
		/// <param name="recipients">object Recipients</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SendMail(object recipients)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients);
			Invoker.Method(this, "SendMail", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx
		/// </summary>
		/// <param name="recipients">object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SendMail(object recipients, object subject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject);
			Invoker.Method(this, "SendMail", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx
		/// </summary>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="priority">optional NetOffice.ExcelApi.Enums.XlPriority Priority = -4143</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SendMailer(object fileFormat, object priority)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileFormat, priority);
			Invoker.Method(this, "SendMailer", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SendMailer()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendMailer", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx
		/// </summary>
		/// <param name="fileFormat">optional object FileFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SendMailer(object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileFormat);
			Invoker.Method(this, "SendMailer", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838177.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="procedure">optional object Procedure</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SetLinkOnData(string name, object procedure)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, procedure);
			Invoker.Method(this, "SetLinkOnData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838177.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SetLinkOnData(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "SetLinkOnData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196695.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196695.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840639.aspx
		/// </summary>
		/// <param name="sharingPassword">optional object SharingPassword</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void UnprotectSharing(object sharingPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sharingPassword);
			Invoker.Method(this, "UnprotectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840639.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void UnprotectSharing()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UnprotectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840979.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void UpdateFromFile()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateFromFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="type">optional object Type</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void UpdateLink(object name, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type);
			Invoker.Method(this, "UpdateLink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void UpdateLink()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateLink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void UpdateLink(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "UpdateLink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx
		/// </summary>
		/// <param name="when">optional object When</param>
		/// <param name="who">optional object Who</param>
		/// <param name="where">optional object Where</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void HighlightChangesOptions(object when, object who, object where)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(when, who, where);
			Invoker.Method(this, "HighlightChangesOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void HighlightChangesOptions()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "HighlightChangesOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx
		/// </summary>
		/// <param name="when">optional object When</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void HighlightChangesOptions(object when)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(when);
			Invoker.Method(this, "HighlightChangesOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx
		/// </summary>
		/// <param name="when">optional object When</param>
		/// <param name="who">optional object Who</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void HighlightChangesOptions(object when, object who)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(when, who);
			Invoker.Method(this, "HighlightChangesOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834662.aspx
		/// </summary>
		/// <param name="days">Int32 Days</param>
		/// <param name="sharingPassword">optional object SharingPassword</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PurgeChangeHistoryNow(Int32 days, object sharingPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(days, sharingPassword);
			Invoker.Method(this, "PurgeChangeHistoryNow", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834662.aspx
		/// </summary>
		/// <param name="days">Int32 Days</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PurgeChangeHistoryNow(Int32 days)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(days);
			Invoker.Method(this, "PurgeChangeHistoryNow", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx
		/// </summary>
		/// <param name="when">optional object When</param>
		/// <param name="who">optional object Who</param>
		/// <param name="where">optional object Where</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void AcceptAllChanges(object when, object who, object where)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(when, who, where);
			Invoker.Method(this, "AcceptAllChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void AcceptAllChanges()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AcceptAllChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx
		/// </summary>
		/// <param name="when">optional object When</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void AcceptAllChanges(object when)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(when);
			Invoker.Method(this, "AcceptAllChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx
		/// </summary>
		/// <param name="when">optional object When</param>
		/// <param name="who">optional object Who</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void AcceptAllChanges(object when, object who)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(when, who);
			Invoker.Method(this, "AcceptAllChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx
		/// </summary>
		/// <param name="when">optional object When</param>
		/// <param name="who">optional object Who</param>
		/// <param name="where">optional object Where</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void RejectAllChanges(object when, object who, object where)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(when, who, where);
			Invoker.Method(this, "RejectAllChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void RejectAllChanges()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RejectAllChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx
		/// </summary>
		/// <param name="when">optional object When</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void RejectAllChanges(object when)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(when);
			Invoker.Method(this, "RejectAllChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx
		/// </summary>
		/// <param name="when">optional object When</param>
		/// <param name="who">optional object Who</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void RejectAllChanges(object when, object who)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(when, who);
			Invoker.Method(this, "RejectAllChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		/// <param name="columnGrand">optional object ColumnGrand</param>
		/// <param name="saveData">optional object SaveData</param>
		/// <param name="hasAutoFormat">optional object HasAutoFormat</param>
		/// <param name="autoPage">optional object AutoPage</param>
		/// <param name="reserved">optional object Reserved</param>
		/// <param name="backgroundQuery">optional object BackgroundQuery</param>
		/// <param name="optimizeCache">optional object OptimizeCache</param>
		/// <param name="pageFieldOrder">optional object PageFieldOrder</param>
		/// <param name="pageFieldWrapCount">optional object PageFieldWrapCount</param>
		/// <param name="readData">optional object ReadData</param>
		/// <param name="connection">optional object Connection</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData, object connection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData, connection);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		/// <param name="columnGrand">optional object ColumnGrand</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		/// <param name="columnGrand">optional object ColumnGrand</param>
		/// <param name="saveData">optional object SaveData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		/// <param name="columnGrand">optional object ColumnGrand</param>
		/// <param name="saveData">optional object SaveData</param>
		/// <param name="hasAutoFormat">optional object HasAutoFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		/// <param name="columnGrand">optional object ColumnGrand</param>
		/// <param name="saveData">optional object SaveData</param>
		/// <param name="hasAutoFormat">optional object HasAutoFormat</param>
		/// <param name="autoPage">optional object AutoPage</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		/// <param name="columnGrand">optional object ColumnGrand</param>
		/// <param name="saveData">optional object SaveData</param>
		/// <param name="hasAutoFormat">optional object HasAutoFormat</param>
		/// <param name="autoPage">optional object AutoPage</param>
		/// <param name="reserved">optional object Reserved</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		/// <param name="columnGrand">optional object ColumnGrand</param>
		/// <param name="saveData">optional object SaveData</param>
		/// <param name="hasAutoFormat">optional object HasAutoFormat</param>
		/// <param name="autoPage">optional object AutoPage</param>
		/// <param name="reserved">optional object Reserved</param>
		/// <param name="backgroundQuery">optional object BackgroundQuery</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		/// <param name="columnGrand">optional object ColumnGrand</param>
		/// <param name="saveData">optional object SaveData</param>
		/// <param name="hasAutoFormat">optional object HasAutoFormat</param>
		/// <param name="autoPage">optional object AutoPage</param>
		/// <param name="reserved">optional object Reserved</param>
		/// <param name="backgroundQuery">optional object BackgroundQuery</param>
		/// <param name="optimizeCache">optional object OptimizeCache</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		/// <param name="columnGrand">optional object ColumnGrand</param>
		/// <param name="saveData">optional object SaveData</param>
		/// <param name="hasAutoFormat">optional object HasAutoFormat</param>
		/// <param name="autoPage">optional object AutoPage</param>
		/// <param name="reserved">optional object Reserved</param>
		/// <param name="backgroundQuery">optional object BackgroundQuery</param>
		/// <param name="optimizeCache">optional object OptimizeCache</param>
		/// <param name="pageFieldOrder">optional object PageFieldOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		/// <param name="columnGrand">optional object ColumnGrand</param>
		/// <param name="saveData">optional object SaveData</param>
		/// <param name="hasAutoFormat">optional object HasAutoFormat</param>
		/// <param name="autoPage">optional object AutoPage</param>
		/// <param name="reserved">optional object Reserved</param>
		/// <param name="backgroundQuery">optional object BackgroundQuery</param>
		/// <param name="optimizeCache">optional object OptimizeCache</param>
		/// <param name="pageFieldOrder">optional object PageFieldOrder</param>
		/// <param name="pageFieldWrapCount">optional object PageFieldWrapCount</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sourceType">optional object SourceType</param>
		/// <param name="sourceData">optional object SourceData</param>
		/// <param name="tableDestination">optional object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="rowGrand">optional object RowGrand</param>
		/// <param name="columnGrand">optional object ColumnGrand</param>
		/// <param name="saveData">optional object SaveData</param>
		/// <param name="hasAutoFormat">optional object HasAutoFormat</param>
		/// <param name="autoPage">optional object AutoPage</param>
		/// <param name="reserved">optional object Reserved</param>
		/// <param name="backgroundQuery">optional object BackgroundQuery</param>
		/// <param name="optimizeCache">optional object OptimizeCache</param>
		/// <param name="pageFieldOrder">optional object PageFieldOrder</param>
		/// <param name="pageFieldWrapCount">optional object PageFieldWrapCount</param>
		/// <param name="readData">optional object ReadData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData);
			Invoker.Method(this, "PivotTableWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194697.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ResetColors()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ResetColors", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="newWindow">optional object NewWindow</param>
		/// <param name="addHistory">optional object AddHistory</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		/// <param name="method">optional object Method</param>
		/// <param name="headerInfo">optional object HeaderInfo</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="newWindow">optional object NewWindow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="newWindow">optional object NewWindow</param>
		/// <param name="addHistory">optional object AddHistory</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="newWindow">optional object NewWindow</param>
		/// <param name="addHistory">optional object AddHistory</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional object SubAddress</param>
		/// <param name="newWindow">optional object NewWindow</param>
		/// <param name="addHistory">optional object AddHistory</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		/// <param name="method">optional object Method</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194282.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void AddToFavorites()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddToFavorites", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="prToFileName">optional object PrToFileName</param>
		/// <param name="ignorePrintAreas">optional object IgnorePrintAreas</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName, object ignorePrintAreas)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName, ignorePrintAreas);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195831.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void WebPagePreview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "WebPagePreview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839234.aspx
		/// </summary>
		/// <param name="encoding">NetOffice.OfficeApi.Enums.MsoEncoding Encoding</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(encoding);
			Invoker.Method(this, "ReloadAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9
		/// 
		/// </summary>
		/// <param name="unused">Int32 unused</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9)]
		public void Dummy1(Int32 unused)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unused);
			Invoker.Method(this, "Dummy1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="s">string s</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void sblt(string s)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(s);
			Invoker.Method(this, "sblt", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="structure">optional object Structure</param>
		/// <param name="windows">optional object Windows</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _Protect(object password, object structure, object windows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, structure, windows);
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
		/// <param name="structure">optional object Structure</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _Protect(object password, object structure)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, structure);
			Invoker.Method(this, "_Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object ConflictResolution</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="textCodepage">optional object TextCodepage</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object ConflictResolution</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object ConflictResolution</param>
		/// <param name="addToMru">optional object AddToMru</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="fileFormat">optional object FileFormat</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object ConflictResolution</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="textCodepage">optional object TextCodepage</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage);
			Invoker.Method(this, "_SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="calcid">Int32 calcid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void Dummy17(Int32 calcid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(calcid);
			Invoker.Method(this, "Dummy17", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194915.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlLinkType Type</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void BreakLink(string name, NetOffice.ExcelApi.Enums.XlLinkType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type);
			Invoker.Method(this, "BreakLink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void Dummy16()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy16", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional object MakePublic</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments, object makePublic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void CheckIn()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="comments">optional object Comments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194456.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool CanCheckIn()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CanCheckIn", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		/// <param name="showMessage">optional object ShowMessage</param>
		/// <param name="includeAttachment">optional object IncludeAttachment</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage, object includeAttachment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage, includeAttachment);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SendForReview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SendForReview(object recipients)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		/// <param name="showMessage">optional object ShowMessage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821626.aspx
		/// </summary>
		/// <param name="showMessage">optional object ShowMessage</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void ReplyWithChanges(object showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showMessage);
			Invoker.Method(this, "ReplyWithChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821626.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void ReplyWithChanges()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReplyWithChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839207.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void EndReview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EndReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx
		/// </summary>
		/// <param name="passwordEncryptionProvider">optional object PasswordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">optional object PasswordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">optional object PasswordEncryptionKeyLength</param>
		/// <param name="passwordEncryptionFileProperties">optional object PasswordEncryptionFileProperties</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm, object passwordEncryptionKeyLength, object passwordEncryptionFileProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength, passwordEncryptionFileProperties);
			Invoker.Method(this, "SetPasswordEncryptionOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SetPasswordEncryptionOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx
		/// </summary>
		/// <param name="passwordEncryptionProvider">optional object PasswordEncryptionProvider</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(object passwordEncryptionProvider)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(passwordEncryptionProvider);
			Invoker.Method(this, "SetPasswordEncryptionOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx
		/// </summary>
		/// <param name="passwordEncryptionProvider">optional object PasswordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">optional object PasswordEncryptionAlgorithm</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(passwordEncryptionProvider, passwordEncryptionAlgorithm);
			Invoker.Method(this, "SetPasswordEncryptionOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx
		/// </summary>
		/// <param name="passwordEncryptionProvider">optional object PasswordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">optional object PasswordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">optional object PasswordEncryptionKeyLength</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm, object passwordEncryptionKeyLength)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength);
			Invoker.Method(this, "SetPasswordEncryptionOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void RecheckSmartTags()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RecheckSmartTags", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		/// <param name="showMessage">optional object ShowMessage</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject, object showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public void SendFaxOverInternet()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx
		/// </summary>
		/// <param name="recipients">optional object Recipients</param>
		/// <param name="subject">optional object Subject</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx
		/// </summary>
		/// <param name="url">string Url</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap ImportMap</param>
		/// <param name="overwrite">optional object Overwrite</param>
		/// <param name="destination">optional object Destination</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap, object overwrite, object destination)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false,false);
			importMap = null;
			object[] paramsArray = Invoker.ValidateParamsArray(url, importMap, overwrite, destination);
			object returnItem = Invoker.MethodReturn(this, "XmlImport", paramsArray);
			importMap = (NetOffice.ExcelApi.XmlMap)paramsArray[1];
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx
		/// </summary>
		/// <param name="url">string Url</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap ImportMap</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			importMap = null;
			object[] paramsArray = Invoker.ValidateParamsArray(url, importMap);
			object returnItem = Invoker.MethodReturn(this, "XmlImport", paramsArray);
			importMap = (NetOffice.ExcelApi.XmlMap)paramsArray[1];
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx
		/// </summary>
		/// <param name="url">string Url</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap ImportMap</param>
		/// <param name="overwrite">optional object Overwrite</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap, object overwrite)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false);
			importMap = null;
			object[] paramsArray = Invoker.ValidateParamsArray(url, importMap, overwrite);
			object returnItem = Invoker.MethodReturn(this, "XmlImport", paramsArray);
			importMap = (NetOffice.ExcelApi.XmlMap)paramsArray[1];
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx
		/// </summary>
		/// <param name="data">string Data</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap ImportMap</param>
		/// <param name="overwrite">optional object Overwrite</param>
		/// <param name="destination">optional object Destination</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap, object overwrite, object destination)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false,false);
			importMap = null;
			object[] paramsArray = Invoker.ValidateParamsArray(data, importMap, overwrite, destination);
			object returnItem = Invoker.MethodReturn(this, "XmlImportXml", paramsArray);
			importMap = (NetOffice.ExcelApi.XmlMap)paramsArray[1];
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx
		/// </summary>
		/// <param name="data">string Data</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap ImportMap</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			importMap = null;
			object[] paramsArray = Invoker.ValidateParamsArray(data, importMap);
			object returnItem = Invoker.MethodReturn(this, "XmlImportXml", paramsArray);
			importMap = (NetOffice.ExcelApi.XmlMap)paramsArray[1];
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx
		/// </summary>
		/// <param name="data">string Data</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap ImportMap</param>
		/// <param name="overwrite">optional object Overwrite</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap, object overwrite)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false);
			importMap = null;
			object[] paramsArray = Invoker.ValidateParamsArray(data, importMap, overwrite);
			object returnItem = Invoker.MethodReturn(this, "XmlImportXml", paramsArray);
			importMap = (NetOffice.ExcelApi.XmlMap)paramsArray[1];
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834616.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="map">NetOffice.ExcelApi.XmlMap Map</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public void SaveAsXMLData(string filename, NetOffice.ExcelApi.XmlMap map)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, map);
			Invoker.Method(this, "SaveAsXMLData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196845.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public void ToggleFormsDesign()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ToggleFormsDesign", paramsArray);
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
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		/// <param name="sharingPassword">optional object SharingPassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, password, writeResPassword, readOnlyRecommended, createBackup, sharingPassword);
			Invoker.Method(this, "_ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void _ProtectSharing()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "_ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="password">optional object Password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, password);
			Invoker.Method(this, "_ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename, object password, object writeResPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, password, writeResPassword);
			Invoker.Method(this, "_ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, password, writeResPassword, readOnlyRecommended);
			Invoker.Method(this, "_ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="readOnlyRecommended">optional object ReadOnlyRecommended</param>
		/// <param name="createBackup">optional object CreateBackup</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, password, writeResPassword, readOnlyRecommended, createBackup);
			Invoker.Method(this, "_ProtectSharing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840327.aspx
		/// </summary>
		/// <param name="removeDocInfoType">NetOffice.ExcelApi.Enums.XlRemoveDocInfoType RemoveDocInfoType</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void RemoveDocumentInformation(NetOffice.ExcelApi.Enums.XlRemoveDocInfoType removeDocInfoType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(removeDocInfoType);
			Invoker.Method(this, "RemoveDocumentInformation", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional object MakePublic</param>
		/// <param name="versionType">optional object VersionType</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic, versionType);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void CheckInWithVersion()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="comments">optional object Comments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional object MakePublic</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838567.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void LockServerFile()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "LockServerFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835507.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetWorkflowTasks", paramsArray);
			NetOffice.OfficeApi.WorkflowTasks newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.WorkflowTasks.LateBindingApiWrapperType) as NetOffice.OfficeApi.WorkflowTasks;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837818.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetWorkflowTemplates", paramsArray);
			NetOffice.OfficeApi.WorkflowTemplates newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.WorkflowTemplates.LateBindingApiWrapperType) as NetOffice.OfficeApi.WorkflowTemplates;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194014.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void ApplyTheme(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "ApplyTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820742.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void EnableConnections()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EnableConnections", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx
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
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void Dummy26()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy26", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void Dummy27()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy27", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}