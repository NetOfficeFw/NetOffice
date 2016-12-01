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
	/// _Application
	///</summary>
	public class _Application_ : COMObject
	{
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Application_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application_() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193687.aspx
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Caller(object index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Caller", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193687.aspx
		/// Alias for get_Caller
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Caller(object index)
		{
			return get_Caller(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836533.aspx
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_ClipboardFormats(object index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "ClipboardFormats", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836533.aspx
		/// Alias for get_ClipboardFormats
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ClipboardFormats(object index)
		{
			return get_ClipboardFormats(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// <param name="index2">optional object Index2</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834406.aspx
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_FileConverters(object index1, object index2)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index1, index2);
			object returnItem = Invoker.PropertyGet(this, "FileConverters", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834406.aspx
		/// Alias for get_FileConverters
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// <param name="index2">optional object Index2</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object FileConverters(object index1, object index2)
		{
			return get_FileConverters(index1, index2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834406.aspx
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_FileConverters(object index1)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index1);
			object returnItem = Invoker.PropertyGet(this, "FileConverters", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834406.aspx
		/// Alias for get_FileConverters
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object FileConverters(object index1)
		{
			return get_FileConverters(index1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840213.aspx
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_International(object index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "International", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840213.aspx
		/// Alias for get_International
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object International(object index)
		{
			return get_International(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836223.aspx
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_PreviousSelections(object index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "PreviousSelections", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836223.aspx
		/// Alias for get_PreviousSelections
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PreviousSelections(object index)
		{
			return get_PreviousSelections(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// <param name="index2">optional object Index2</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839011.aspx
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_RegisteredFunctions(object index1, object index2)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index1, index2);
			object returnItem = Invoker.PropertyGet(this, "RegisteredFunctions", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839011.aspx
		/// Alias for get_RegisteredFunctions
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// <param name="index2">optional object Index2</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object RegisteredFunctions(object index1, object index2)
		{
			return get_RegisteredFunctions(index1, index2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839011.aspx
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_RegisteredFunctions(object index1)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index1);
			object returnItem = Invoker.PropertyGet(this, "RegisteredFunctions", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839011.aspx
		/// Alias for get_RegisteredFunctions
		/// </summary>
		/// <param name="index1">optional object Index1</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object RegisteredFunctions(object index1)
		{
			return get_RegisteredFunctions(index1);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// DispatchInterface _Application 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Application : _Application_
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
                    _type = typeof(_Application);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Application(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193029.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836139.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840249.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.ExcelApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Application.LateBindingApiWrapperType) as NetOffice.ExcelApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834673.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range ActiveCell
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveCell", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196586.aspx
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
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.DialogSheet ActiveDialog
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveDialog", paramsArray);
				NetOffice.ExcelApi.DialogSheet newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.DialogSheet.LateBindingApiWrapperType) as NetOffice.ExcelApi.DialogSheet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.MenuBar ActiveMenuBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveMenuBar", paramsArray);
				NetOffice.ExcelApi.MenuBar newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.MenuBar.LateBindingApiWrapperType) as NetOffice.ExcelApi.MenuBar;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822927.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string ActivePrinter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActivePrinter", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ActivePrinter", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822753.aspx
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
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835855.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Window ActiveWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveWindow", paramsArray);
				NetOffice.ExcelApi.Window newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Window.LateBindingApiWrapperType) as NetOffice.ExcelApi.Window;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821871.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook ActiveWorkbook
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveWorkbook", paramsArray);
				NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193312.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.AddIns AddIns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AddIns", paramsArray);
				NetOffice.ExcelApi.AddIns newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.AddIns.LateBindingApiWrapperType) as NetOffice.ExcelApi.AddIns;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Assistant", paramsArray);
				NetOffice.OfficeApi.Assistant newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType) as NetOffice.OfficeApi.Assistant;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836446.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Cells
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839597.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195212.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range Columns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Columns", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838002.aspx
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
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836739.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 DDEAppReturnCode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DDEAppReturnCode", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
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
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.MenuBars MenuBars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MenuBars", paramsArray);
				NetOffice.ExcelApi.MenuBars newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.MenuBars.LateBindingApiWrapperType) as NetOffice.ExcelApi.MenuBars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Modules Modules
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Modules", paramsArray);
				NetOffice.ExcelApi.Modules newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Modules.LateBindingApiWrapperType) as NetOffice.ExcelApi.Modules;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841251.aspx
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
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841254.aspx
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		/// <param name="cell2">optional object Cell2</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Range(object cell1, object cell2)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(cell1, cell2);
			object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841254.aspx
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		/// <param name="cell2">optional object Cell2</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Range(object cell1, object cell2)
		{
			return get_Range(cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841254.aspx
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Range(object cell1)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(cell1);
			object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841254.aspx
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Range(object cell1)
		{
			return get_Range(cell1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197831.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range Rows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Rows", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840834.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822920.aspx
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
		/// Get
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Menu get_ShortcutMenus(Int32 index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "ShortcutMenus", paramsArray);
			NetOffice.ExcelApi.Menu newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Menu.LateBindingApiWrapperType) as NetOffice.ExcelApi.Menu;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ShortcutMenus
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Menu ShortcutMenus(Int32 index)
		{
			return get_ShortcutMenus(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193227.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook ThisWorkbook
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ThisWorkbook", paramsArray);
				NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Toolbars Toolbars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Toolbars", paramsArray);
				NetOffice.ExcelApi.Toolbars newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Toolbars.LateBindingApiWrapperType) as NetOffice.ExcelApi.Toolbars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193325.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820765.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbooks Workbooks
		{
			get
            {
                object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Workbooks", paramsArray);
				NetOffice.ExcelApi.Workbooks newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Workbooks.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbooks;
                return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841212.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.WorksheetFunction WorksheetFunction
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WorksheetFunction", paramsArray);
				NetOffice.ExcelApi.WorksheetFunction newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.WorksheetFunction.LateBindingApiWrapperType) as NetOffice.ExcelApi.WorksheetFunction;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840672.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196733.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839420.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823129.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool AlertBeforeOverwriting
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AlertBeforeOverwriting", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AlertBeforeOverwriting", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836135.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string AltStartupPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AltStartupPath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AltStartupPath", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194812.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool AskToUpdateLinks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AskToUpdateLinks", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AskToUpdateLinks", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841180.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool EnableAnimations
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableAnimations", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableAnimations", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840200.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.AutoCorrect AutoCorrect
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoCorrect", paramsArray);
				NetOffice.ExcelApi.AutoCorrect newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.AutoCorrect.LateBindingApiWrapperType) as NetOffice.ExcelApi.AutoCorrect;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839811.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Build
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Build", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194157.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool CalculateBeforeSave
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CalculateBeforeSave", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CalculateBeforeSave", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821260.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCalculation Calculation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Calculation", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCalculation)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Calculation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193687.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Caller
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Caller", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198163.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool CanPlaySounds
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CanPlaySounds", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837138.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool CanRecordSounds
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CanRecordSounds", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821801.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string Caption
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Caption", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Caption", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839786.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool CellDragAndDrop
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CellDragAndDrop", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CellDragAndDrop", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836533.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ClipboardFormats
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ClipboardFormats", paramsArray);
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
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194368.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayClipboardWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayClipboardWindow", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayClipboardWindow", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool ColorButtons
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColorButtons", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ColorButtons", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193318.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCommandUnderlines CommandUnderlines
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandUnderlines", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCommandUnderlines)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CommandUnderlines", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835900.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ConstrainNumeric
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConstrainNumeric", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ConstrainNumeric", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835243.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool CopyObjectsWithCells
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CopyObjectsWithCells", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CopyObjectsWithCells", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198335.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlMousePointer Cursor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Cursor", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlMousePointer)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Cursor", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836479.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 CustomListCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomListCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839532.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCutCopyMode CutCopyMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CutCopyMode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCutCopyMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CutCopyMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194983.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 DataEntryMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataEntryMode", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DataEntryMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string _Default
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "_Default", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835839.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string DefaultFilePath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultFilePath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultFilePath", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193689.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Dialogs Dialogs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dialogs", paramsArray);
				NetOffice.ExcelApi.Dialogs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Dialogs.LateBindingApiWrapperType) as NetOffice.ExcelApi.Dialogs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839782.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayAlerts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayAlerts", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayAlerts", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837410.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayFormulaBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayFormulaBar", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayFormulaBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838060.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayFullScreen
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayFullScreen", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayFullScreen", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836230.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayNoteIndicator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayNoteIndicator", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayNoteIndicator", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835215.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCommentDisplayMode DisplayCommentIndicator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayCommentIndicator", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCommentDisplayMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayCommentIndicator", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838763.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayExcel4Menus
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayExcel4Menus", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayExcel4Menus", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192941.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayRecentFiles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayRecentFiles", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayRecentFiles", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837572.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayScrollBars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayScrollBars", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayScrollBars", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838599.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayStatusBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayStatusBar", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayStatusBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840396.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool EditDirectlyInCell
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EditDirectlyInCell", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EditDirectlyInCell", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840522.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool EnableAutoComplete
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableAutoComplete", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableAutoComplete", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834623.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlEnableCancelKey EnableCancelKey
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableCancelKey", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlEnableCancelKey)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableCancelKey", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834964.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool EnableSound
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableSound", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableSound", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool EnableTipWizard
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableTipWizard", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableTipWizard", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834406.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object FileConverters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileConverters", paramsArray);
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
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FileSearch FileSearch
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileSearch", paramsArray);
				NetOffice.OfficeApi.FileSearch newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.FileSearch.LateBindingApiWrapperType) as NetOffice.OfficeApi.FileSearch;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.IFind FileFind
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileFind", paramsArray);
				NetOffice.OfficeApi.IFind newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IFind.LateBindingApiWrapperType) as NetOffice.OfficeApi.IFind;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197833.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool FixedDecimal
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FixedDecimal", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FixedDecimal", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840166.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 FixedDecimalPlaces
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FixedDecimalPlaces", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FixedDecimalPlaces", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195524.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Double Height
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Height", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Height", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836184.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool IgnoreRemoteRequests
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IgnoreRemoteRequests", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IgnoreRemoteRequests", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841248.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Interactive
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Interactive", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Interactive", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840213.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object International
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "International", paramsArray);
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
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820731.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Iteration
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Iteration", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Iteration", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool LargeButtons
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LargeButtons", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LargeButtons", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834367.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Double Left
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Left", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Left", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834314.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string LibraryPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LibraryPath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197561.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object MailSession
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailSession", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840051.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlMailSystem MailSystem
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailSystem", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlMailSystem)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836180.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool MathCoprocessorAvailable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MathCoprocessorAvailable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820917.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Double MaxChange
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MaxChange", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MaxChange", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834974.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 MaxIterations
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MaxIterations", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MaxIterations", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 MemoryFree
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MemoryFree", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 MemoryTotal
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MemoryTotal", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 MemoryUsed
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MemoryUsed", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838014.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool MouseAvailable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MouseAvailable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836785.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool MoveAfterReturn
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MoveAfterReturn", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MoveAfterReturn", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838636.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlDirection MoveAfterReturnDirection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MoveAfterReturnDirection", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlDirection)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MoveAfterReturnDirection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837432.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.RecentFiles RecentFiles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RecentFiles", paramsArray);
				NetOffice.ExcelApi.RecentFiles newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.RecentFiles.LateBindingApiWrapperType) as NetOffice.ExcelApi.RecentFiles;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841101.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197616.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string NetworkTemplatesPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NetworkTemplatesPath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197725.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ODBCErrors ODBCErrors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ODBCErrors", paramsArray);
				NetOffice.ExcelApi.ODBCErrors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ODBCErrors.LateBindingApiWrapperType) as NetOffice.ExcelApi.ODBCErrors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835926.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 ODBCTimeout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ODBCTimeout", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ODBCTimeout", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnCalculate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OnCalculate", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OnCalculate", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OnData", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OnData", paramsArray);
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
		public string OnEntry
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OnEntry", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OnEntry", paramsArray);
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
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823040.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string OnWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OnWindow", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OnWindow", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837365.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string OperatingSystem
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OperatingSystem", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197291.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string OrganizationName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OrganizationName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193842.aspx
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
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820973.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string PathSeparator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PathSeparator", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836223.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PreviousSelections
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PreviousSelections", paramsArray);
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
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840093.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool PivotTableSelection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PivotTableSelection", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PivotTableSelection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822374.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool PromptForSummaryInfo
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PromptForSummaryInfo", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PromptForSummaryInfo", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821936.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool RecordRelative
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RecordRelative", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835250.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReferenceStyle", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlReferenceStyle)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ReferenceStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839011.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object RegisteredFunctions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RegisteredFunctions", paramsArray);
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
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193630.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool RollZoom
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RollZoom", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RollZoom", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193498.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ScreenUpdating
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScreenUpdating", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ScreenUpdating", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840164.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetsInNewWorkbook
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SheetsInNewWorkbook", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SheetsInNewWorkbook", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837068.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ShowChartTipNames
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowChartTipNames", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowChartTipNames", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835275.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ShowChartTipValues
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowChartTipValues", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowChartTipValues", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822527.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string StandardFont
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StandardFont", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "StandardFont", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196551.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Double StandardFontSize
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StandardFontSize", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "StandardFontSize", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193231.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string StartupPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StartupPath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835916.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object StatusBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StatusBar", paramsArray);
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
				Invoker.PropertySet(this, "StatusBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195932.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string TemplatesPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TemplatesPath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822844.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ShowToolTips
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowToolTips", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowToolTips", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196150.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Double Top
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Top", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Top", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838398.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlFileFormat DefaultSaveFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultSaveFormat", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlFileFormat)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultSaveFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197155.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string TransitionMenuKey
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TransitionMenuKey", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TransitionMenuKey", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835846.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 TransitionMenuKeyAction
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TransitionMenuKeyAction", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TransitionMenuKeyAction", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195370.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool TransitionNavigKeys
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TransitionNavigKeys", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TransitionNavigKeys", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820823.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Double UsableHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UsableHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838216.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Double UsableWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UsableWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841219.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822584.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string UserName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UserName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UserName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195664.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string Value
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Value", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840360.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBE", paramsArray);
				NetOffice.VBIDEApi.VBE newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBE;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193301.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string Version
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Version", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198119.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Visible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Visible", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Visible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840678.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Double Width
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Width", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Width", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834350.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool WindowsForPens
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WindowsForPens", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840986.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlWindowState WindowState
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WindowState", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlWindowState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WindowState", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 UILanguage
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UILanguage", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UILanguage", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196339.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 DefaultSheetDirection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultSheetDirection", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultSheetDirection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198001.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 CursorMovement
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CursorMovement", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CursorMovement", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193043.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ControlCharacters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ControlCharacters", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ControlCharacters", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821508.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool EnableEvents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableEvents", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableEvents", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayInfoWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayInfoWindow", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayInfoWindow", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838041.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ExtendList
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ExtendList", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ExtendList", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193551.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.OLEDBErrors OLEDBErrors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OLEDBErrors", paramsArray);
				NetOffice.ExcelApi.OLEDBErrors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.OLEDBErrors.LateBindingApiWrapperType) as NetOffice.ExcelApi.OLEDBErrors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839557.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.COMAddIns COMAddIns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "COMAddIns", paramsArray);
				NetOffice.OfficeApi.COMAddIns newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType) as NetOffice.OfficeApi.COMAddIns;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198342.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.DefaultWebOptions DefaultWebOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultWebOptions", paramsArray);
				NetOffice.ExcelApi.DefaultWebOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.DefaultWebOptions.LateBindingApiWrapperType) as NetOffice.ExcelApi.DefaultWebOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821592.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string ProductCode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProductCode", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197803.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string UserLibraryPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UserLibraryPath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834750.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool AutoPercentEntry
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoPercentEntry", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutoPercentEntry", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821852.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.LanguageSettings LanguageSettings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LanguageSettings", paramsArray);
				NetOffice.OfficeApi.LanguageSettings newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.LanguageSettings.LateBindingApiWrapperType) as NetOffice.OfficeApi.LanguageSettings;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Dummy101
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dummy101", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.AnswerWizard AnswerWizard
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AnswerWizard", paramsArray);
				NetOffice.OfficeApi.AnswerWizard newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.AnswerWizard.LateBindingApiWrapperType) as NetOffice.OfficeApi.AnswerWizard;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193990.aspx
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
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool ShowWindowsInTaskbar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowWindowsInTaskbar", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowWindowsInTaskbar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193655.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FeatureInstall", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoFeatureInstall)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FeatureInstall", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197917.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool Ready
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Ready", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838023.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CellFormat FindFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FindFormat", paramsArray);
				NetOffice.ExcelApi.CellFormat newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.CellFormat.LateBindingApiWrapperType) as NetOffice.ExcelApi.CellFormat;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FindFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840036.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CellFormat ReplaceFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReplaceFormat", paramsArray);
				NetOffice.ExcelApi.CellFormat newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.CellFormat.LateBindingApiWrapperType) as NetOffice.ExcelApi.CellFormat;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ReplaceFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838590.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.UsedObjects UsedObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UsedObjects", paramsArray);
				NetOffice.ExcelApi.UsedObjects newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.UsedObjects.LateBindingApiWrapperType) as NetOffice.ExcelApi.UsedObjects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196047.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCalculationState CalculationState
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CalculationState", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCalculationState)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194059.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCalculationInterruptKey CalculationInterruptKey
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CalculationInterruptKey", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCalculationInterruptKey)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CalculationInterruptKey", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197781.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Watches Watches
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Watches", paramsArray);
				NetOffice.ExcelApi.Watches newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Watches.LateBindingApiWrapperType) as NetOffice.ExcelApi.Watches;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839155.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool DisplayFunctionToolTips
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayFunctionToolTips", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayFunctionToolTips", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837822.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutomationSecurity", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoAutomationSecurity)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutomationSecurity", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836226.aspx
		/// </summary>
		/// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(fileDialogType);
			object returnItem = Invoker.PropertyGet(this, "FileDialog", paramsArray);
			NetOffice.OfficeApi.FileDialog newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.FileDialog.LateBindingApiWrapperType) as NetOffice.OfficeApi.FileDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836226.aspx
		/// Alias for get_FileDialog
		/// </summary>
		/// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType)
		{
			return get_FileDialog(fileDialogType);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839816.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool DisplayPasteOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayPasteOptions", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayPasteOptions", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834919.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool DisplayInsertOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayInsertOptions", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayInsertOptions", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834973.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool GenerateGetPivotData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GenerateGetPivotData", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GenerateGetPivotData", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838418.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.AutoRecover AutoRecover
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoRecover", paramsArray);
				NetOffice.ExcelApi.AutoRecover newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.AutoRecover.LateBindingApiWrapperType) as NetOffice.ExcelApi.AutoRecover;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840629.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 Hwnd
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hwnd", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197539.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 Hinstance
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hinstance", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196618.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ErrorCheckingOptions ErrorCheckingOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ErrorCheckingOptions", paramsArray);
				NetOffice.ExcelApi.ErrorCheckingOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ErrorCheckingOptions.LateBindingApiWrapperType) as NetOffice.ExcelApi.ErrorCheckingOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836132.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceHyperlinks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoFormatAsYouTypeReplaceHyperlinks", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutoFormatAsYouTypeReplaceHyperlinks", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.SmartTagRecognizers SmartTagRecognizers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartTagRecognizers", paramsArray);
				NetOffice.ExcelApi.SmartTagRecognizers newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.SmartTagRecognizers.LateBindingApiWrapperType) as NetOffice.ExcelApi.SmartTagRecognizers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196842.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.NewFile NewWorkbook
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NewWorkbook", paramsArray);
				NetOffice.OfficeApi.NewFile newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.NewFile.LateBindingApiWrapperType) as NetOffice.OfficeApi.NewFile;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838794.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.SpellingOptions SpellingOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SpellingOptions", paramsArray);
				NetOffice.ExcelApi.SpellingOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.SpellingOptions.LateBindingApiWrapperType) as NetOffice.ExcelApi.SpellingOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836468.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Speech Speech
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Speech", paramsArray);
				NetOffice.ExcelApi.Speech newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Speech.LateBindingApiWrapperType) as NetOffice.ExcelApi.Speech;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838652.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool MapPaperSize
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MapPaperSize", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MapPaperSize", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835836.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool ShowStartupDialog
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowStartupDialog", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowStartupDialog", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195207.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public string DecimalSeparator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DecimalSeparator", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DecimalSeparator", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839793.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public string ThousandsSeparator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ThousandsSeparator", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ThousandsSeparator", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840692.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool UseSystemSeparators
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UseSystemSeparators", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UseSystemSeparators", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834969.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range ThisCell
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ThisCell", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840124.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.RTD RTD
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RTD", paramsArray);
				NetOffice.ExcelApi.RTD newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.RTD.LateBindingApiWrapperType) as NetOffice.ExcelApi.RTD;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196875.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public bool DisplayDocumentActionTaskPane
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayDocumentActionTaskPane", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayDocumentActionTaskPane", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841030.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public bool ArbitraryXMLSupportAvailable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ArbitraryXMLSupportAvailable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196051.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 MeasurementUnit
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MeasurementUnit", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MeasurementUnit", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839447.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool ShowSelectionFloaties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowSelectionFloaties", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowSelectionFloaties", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835572.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool ShowMenuFloaties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowMenuFloaties", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowMenuFloaties", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839968.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool ShowDevTools
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowDevTools", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowDevTools", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197492.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool EnableLivePreview
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableLivePreview", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableLivePreview", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192972.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool DisplayDocumentInformationPanel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayDocumentInformationPanel", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayDocumentInformationPanel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841106.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool AlwaysUseClearType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AlwaysUseClearType", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AlwaysUseClearType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838767.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool WarnOnFunctionNameConflict
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WarnOnFunctionNameConflict", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WarnOnFunctionNameConflict", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841264.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 FormulaBarHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormulaBarHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormulaBarHeight", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838451.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool DisplayFormulaAutoComplete
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayFormulaAutoComplete", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayFormulaAutoComplete", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196417.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlGenerateTableRefs GenerateTableRefs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GenerateTableRefs", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlGenerateTableRefs)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GenerateTableRefs", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838606.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.IAssistance Assistance
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Assistance", paramsArray);
				NetOffice.OfficeApi.IAssistance newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType) as NetOffice.OfficeApi.IAssistance;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839005.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool EnableLargeOperationAlert
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableLargeOperationAlert", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableLargeOperationAlert", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841027.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 LargeOperationCellThousandCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LargeOperationCellThousandCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LargeOperationCellThousandCount", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195064.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool DeferAsyncQueries
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DeferAsyncQueries", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DeferAsyncQueries", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835198.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.MultiThreadedCalculation MultiThreadedCalculation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MultiThreadedCalculation", paramsArray);
				NetOffice.ExcelApi.MultiThreadedCalculation newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.MultiThreadedCalculation.LateBindingApiWrapperType) as NetOffice.ExcelApi.MultiThreadedCalculation;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837115.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 ActiveEncryptionSession
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveEncryptionSession", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822842.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public bool HighQualityModeForGraphics
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HighQualityModeForGraphics", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HighQualityModeForGraphics", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194700.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.FileExportConverters FileExportConverters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileExportConverters", paramsArray);
				NetOffice.ExcelApi.FileExportConverters newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.FileExportConverters.LateBindingApiWrapperType) as NetOffice.ExcelApi.FileExportConverters;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192963.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtLayouts SmartArtLayouts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartArtLayouts", paramsArray);
				NetOffice.OfficeApi.SmartArtLayouts newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartArtLayouts.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartArtLayouts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839548.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtQuickStyles SmartArtQuickStyles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartArtQuickStyles", paramsArray);
				NetOffice.OfficeApi.SmartArtQuickStyles newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartArtQuickStyles.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartArtQuickStyles;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194216.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtColors SmartArtColors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartArtColors", paramsArray);
				NetOffice.OfficeApi.SmartArtColors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartArtColors.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartArtColors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197196.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.AddIns2 AddIns2
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AddIns2", paramsArray);
				NetOffice.ExcelApi.AddIns2 newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.AddIns2.LateBindingApiWrapperType) as NetOffice.ExcelApi.AddIns2;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835544.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public bool PrintCommunication
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintCommunication", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintCommunication", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836816.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public bool UseClusterConnector
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UseClusterConnector", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UseClusterConnector", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820828.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public string ClusterConnector
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ClusterConnector", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ClusterConnector", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Quitting
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Quitting", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Dummy22
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dummy22", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Dummy22", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Dummy23
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dummy23", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Dummy23", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193852.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.ProtectedViewWindows ProtectedViewWindows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectedViewWindows", paramsArray);
				NetOffice.ExcelApi.ProtectedViewWindows newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ProtectedViewWindows.LateBindingApiWrapperType) as NetOffice.ExcelApi.ProtectedViewWindows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195068.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.ProtectedViewWindow ActiveProtectedViewWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveProtectedViewWindow", paramsArray);
				NetOffice.ExcelApi.ProtectedViewWindow newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.ExcelApi.ProtectedViewWindow;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839573.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public bool IsSandboxed
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsSandboxed", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public bool SaveISO8601Dates
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SaveISO8601Dates", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SaveISO8601Dates", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841235.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public object HinstancePtr
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HinstancePtr", paramsArray);
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822746.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFileValidationMode FileValidation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileValidation", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoFileValidationMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FileValidation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196973.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlFileValidationPivotMode FileValidationPivot
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileValidationPivot", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlFileValidationPivotMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FileValidationPivot", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227320.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public bool ShowQuickAnalysis
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowQuickAnalysis", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowQuickAnalysis", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231547.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.QuickAnalysis QuickAnalysis
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "QuickAnalysis", paramsArray);
				NetOffice.ExcelApi.QuickAnalysis newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.QuickAnalysis.LateBindingApiWrapperType) as NetOffice.ExcelApi.QuickAnalysis;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230556.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public bool FlashFill
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FlashFill", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FlashFill", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231284.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public bool EnableMacroAnimations
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableMacroAnimations", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableMacroAnimations", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227694.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231743.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public bool FlashFillMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FlashFillMode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FlashFillMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public bool MergeInstances
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MergeInstances", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MergeInstances", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195517.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Calculate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Calculate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194507.aspx
		/// </summary>
		/// <param name="channel">Int32 Channel</param>
		/// <param name="_string">string String</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void DDEExecute(Int32 channel, string _string)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(channel, _string);
			Invoker.Method(this, "DDEExecute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197893.aspx
		/// </summary>
		/// <param name="app">string App</param>
		/// <param name="topic">string Topic</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 DDEInitiate(string app, string topic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(app, topic);
			object returnItem = Invoker.MethodReturn(this, "DDEInitiate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821378.aspx
		/// </summary>
		/// <param name="channel">Int32 Channel</param>
		/// <param name="item">object Item</param>
		/// <param name="data">object Data</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void DDEPoke(Int32 channel, object item, object data)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(channel, item, data);
			Invoker.Method(this, "DDEPoke", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834935.aspx
		/// </summary>
		/// <param name="channel">Int32 Channel</param>
		/// <param name="item">string Item</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DDERequest(Int32 channel, string item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(channel, item);
			object returnItem = Invoker.MethodReturn(this, "DDERequest", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840740.aspx
		/// </summary>
		/// <param name="channel">Int32 Channel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void DDETerminate(Int32 channel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(channel);
			Invoker.Method(this, "DDETerminate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193019.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193589.aspx
		/// </summary>
		/// <param name="_string">string String</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ExecuteExcel4Macro(string _string)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_string);
			object returnItem = Invoker.MethodReturn(this, "ExecuteExcel4Macro", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		/// <param name="arg30">optional object Arg30</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835030.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29);
			object returnItem = Invoker.MethodReturn(this, "Intersect", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		/// <param name="arg30">optional object Arg30</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197132.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		/// <param name="arg30">optional object Arg30</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// <param name="macro">optional object Macro</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29);
			object returnItem = Invoker.MethodReturn(this, "_Run2", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821075.aspx
		/// </summary>
		/// <param name="keys">object Keys</param>
		/// <param name="wait">optional object Wait</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SendKeys(object keys, object wait)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(keys, wait);
			Invoker.Method(this, "SendKeys", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821075.aspx
		/// </summary>
		/// <param name="keys">object Keys</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SendKeys(object keys)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(keys);
			Invoker.Method(this, "SendKeys", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		/// <param name="arg30">optional object Arg30</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834621.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range Arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29);
			object returnItem = Invoker.MethodReturn(this, "Union", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840097.aspx
		/// </summary>
		/// <param name="index">NetOffice.ExcelApi.Enums.XlMSApplication Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ActivateMicrosoftApp(NetOffice.ExcelApi.Enums.XlMSApplication index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "ActivateMicrosoftApp", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="chart">object Chart</param>
		/// <param name="name">string Name</param>
		/// <param name="description">optional object Description</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void AddChartAutoFormat(object chart, string name, object description)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chart, name, description);
			Invoker.Method(this, "AddChartAutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="chart">object Chart</param>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void AddChartAutoFormat(object chart, string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chart, name);
			Invoker.Method(this, "AddChartAutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196224.aspx
		/// </summary>
		/// <param name="listArray">object ListArray</param>
		/// <param name="byRow">optional object ByRow</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void AddCustomList(object listArray, object byRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listArray, byRow);
			Invoker.Method(this, "AddCustomList", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196224.aspx
		/// </summary>
		/// <param name="listArray">object ListArray</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void AddCustomList(object listArray)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listArray);
			Invoker.Method(this, "AddCustomList", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195406.aspx
		/// </summary>
		/// <param name="centimeters">Double Centimeters</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Double CentimetersToPoints(Double centimeters)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(centimeters);
			object returnItem = Invoker.MethodReturn(this, "CentimetersToPoints", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840059.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840059.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840059.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822751.aspx
		/// </summary>
		/// <param name="formula">object Formula</param>
		/// <param name="fromReferenceStyle">NetOffice.ExcelApi.Enums.XlReferenceStyle FromReferenceStyle</param>
		/// <param name="toReferenceStyle">optional object ToReferenceStyle</param>
		/// <param name="toAbsolute">optional object ToAbsolute</param>
		/// <param name="relativeTo">optional object RelativeTo</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ConvertFormula(object formula, NetOffice.ExcelApi.Enums.XlReferenceStyle fromReferenceStyle, object toReferenceStyle, object toAbsolute, object relativeTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formula, fromReferenceStyle, toReferenceStyle, toAbsolute, relativeTo);
			object returnItem = Invoker.MethodReturn(this, "ConvertFormula", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822751.aspx
		/// </summary>
		/// <param name="formula">object Formula</param>
		/// <param name="fromReferenceStyle">NetOffice.ExcelApi.Enums.XlReferenceStyle FromReferenceStyle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ConvertFormula(object formula, NetOffice.ExcelApi.Enums.XlReferenceStyle fromReferenceStyle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formula, fromReferenceStyle);
			object returnItem = Invoker.MethodReturn(this, "ConvertFormula", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822751.aspx
		/// </summary>
		/// <param name="formula">object Formula</param>
		/// <param name="fromReferenceStyle">NetOffice.ExcelApi.Enums.XlReferenceStyle FromReferenceStyle</param>
		/// <param name="toReferenceStyle">optional object ToReferenceStyle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ConvertFormula(object formula, NetOffice.ExcelApi.Enums.XlReferenceStyle fromReferenceStyle, object toReferenceStyle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formula, fromReferenceStyle, toReferenceStyle);
			object returnItem = Invoker.MethodReturn(this, "ConvertFormula", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822751.aspx
		/// </summary>
		/// <param name="formula">object Formula</param>
		/// <param name="fromReferenceStyle">NetOffice.ExcelApi.Enums.XlReferenceStyle FromReferenceStyle</param>
		/// <param name="toReferenceStyle">optional object ToReferenceStyle</param>
		/// <param name="toAbsolute">optional object ToAbsolute</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ConvertFormula(object formula, NetOffice.ExcelApi.Enums.XlReferenceStyle fromReferenceStyle, object toReferenceStyle, object toAbsolute)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formula, fromReferenceStyle, toReferenceStyle, toAbsolute);
			object returnItem = Invoker.MethodReturn(this, "ConvertFormula", paramsArray);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Dummy1()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy1(object arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "Dummy1", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy1(object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1);
			object returnItem = Invoker.MethodReturn(this, "Dummy1", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy1(object arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "Dummy1", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy1(object arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "Dummy1", paramsArray);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Dummy2()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy2(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(this, "Dummy2", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy2(object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1);
			object returnItem = Invoker.MethodReturn(this, "Dummy2", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy2(object arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "Dummy2", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy2(object arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "Dummy2", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy2(object arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "Dummy2", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy2(object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(this, "Dummy2", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy2(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(this, "Dummy2", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy2(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(this, "Dummy2", paramsArray);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Dummy3()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy3", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Dummy4()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy4", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy4(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14);
			object returnItem = Invoker.MethodReturn(this, "Dummy4", paramsArray);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Dummy5()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy5", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy5(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
			object returnItem = Invoker.MethodReturn(this, "Dummy5", paramsArray);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Dummy6()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy6", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Dummy7()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy7", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Dummy8()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy8", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy8(object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1);
			object returnItem = Invoker.MethodReturn(this, "Dummy8", paramsArray);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Dummy9()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy9", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Dummy10()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy10", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg">optional object arg</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool Dummy10(object arg)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg);
			object returnItem = Invoker.MethodReturn(this, "Dummy10", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Dummy11()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy11", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void DeleteChartAutoFormat(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "DeleteChartAutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197264.aspx
		/// </summary>
		/// <param name="listNum">Int32 ListNum</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void DeleteCustomList(Int32 listNum)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listNum);
			Invoker.Method(this, "DeleteCustomList", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194422.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void DoubleClick()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DoubleClick", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _FindFile()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_FindFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196861.aspx
		/// </summary>
		/// <param name="listNum">Int32 ListNum</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetCustomListContents(Int32 listNum)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listNum);
			object returnItem = Invoker.MethodReturn(this, "GetCustomListContents", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838809.aspx
		/// </summary>
		/// <param name="listArray">object ListArray</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 GetCustomListNum(object listArray)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listArray);
			object returnItem = Invoker.MethodReturn(this, "GetCustomListNum", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834966.aspx
		/// </summary>
		/// <param name="fileFilter">optional object FileFilter</param>
		/// <param name="filterIndex">optional object FilterIndex</param>
		/// <param name="title">optional object Title</param>
		/// <param name="buttonText">optional object ButtonText</param>
		/// <param name="multiSelect">optional object MultiSelect</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetOpenFilename(object fileFilter, object filterIndex, object title, object buttonText, object multiSelect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileFilter, filterIndex, title, buttonText, multiSelect);
			object returnItem = Invoker.MethodReturn(this, "GetOpenFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834966.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetOpenFilename()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetOpenFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834966.aspx
		/// </summary>
		/// <param name="fileFilter">optional object FileFilter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetOpenFilename(object fileFilter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileFilter);
			object returnItem = Invoker.MethodReturn(this, "GetOpenFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834966.aspx
		/// </summary>
		/// <param name="fileFilter">optional object FileFilter</param>
		/// <param name="filterIndex">optional object FilterIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetOpenFilename(object fileFilter, object filterIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileFilter, filterIndex);
			object returnItem = Invoker.MethodReturn(this, "GetOpenFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834966.aspx
		/// </summary>
		/// <param name="fileFilter">optional object FileFilter</param>
		/// <param name="filterIndex">optional object FilterIndex</param>
		/// <param name="title">optional object Title</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetOpenFilename(object fileFilter, object filterIndex, object title)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileFilter, filterIndex, title);
			object returnItem = Invoker.MethodReturn(this, "GetOpenFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834966.aspx
		/// </summary>
		/// <param name="fileFilter">optional object FileFilter</param>
		/// <param name="filterIndex">optional object FilterIndex</param>
		/// <param name="title">optional object Title</param>
		/// <param name="buttonText">optional object ButtonText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetOpenFilename(object fileFilter, object filterIndex, object title, object buttonText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileFilter, filterIndex, title, buttonText);
			object returnItem = Invoker.MethodReturn(this, "GetOpenFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195734.aspx
		/// </summary>
		/// <param name="initialFilename">optional object InitialFilename</param>
		/// <param name="fileFilter">optional object FileFilter</param>
		/// <param name="filterIndex">optional object FilterIndex</param>
		/// <param name="title">optional object Title</param>
		/// <param name="buttonText">optional object ButtonText</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetSaveAsFilename(object initialFilename, object fileFilter, object filterIndex, object title, object buttonText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialFilename, fileFilter, filterIndex, title, buttonText);
			object returnItem = Invoker.MethodReturn(this, "GetSaveAsFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195734.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetSaveAsFilename()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetSaveAsFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195734.aspx
		/// </summary>
		/// <param name="initialFilename">optional object InitialFilename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetSaveAsFilename(object initialFilename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialFilename);
			object returnItem = Invoker.MethodReturn(this, "GetSaveAsFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195734.aspx
		/// </summary>
		/// <param name="initialFilename">optional object InitialFilename</param>
		/// <param name="fileFilter">optional object FileFilter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetSaveAsFilename(object initialFilename, object fileFilter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialFilename, fileFilter);
			object returnItem = Invoker.MethodReturn(this, "GetSaveAsFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195734.aspx
		/// </summary>
		/// <param name="initialFilename">optional object InitialFilename</param>
		/// <param name="fileFilter">optional object FileFilter</param>
		/// <param name="filterIndex">optional object FilterIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetSaveAsFilename(object initialFilename, object fileFilter, object filterIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialFilename, fileFilter, filterIndex);
			object returnItem = Invoker.MethodReturn(this, "GetSaveAsFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195734.aspx
		/// </summary>
		/// <param name="initialFilename">optional object InitialFilename</param>
		/// <param name="fileFilter">optional object FileFilter</param>
		/// <param name="filterIndex">optional object FilterIndex</param>
		/// <param name="title">optional object Title</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object GetSaveAsFilename(object initialFilename, object fileFilter, object filterIndex, object title)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialFilename, fileFilter, filterIndex, title);
			object returnItem = Invoker.MethodReturn(this, "GetSaveAsFilename", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839232.aspx
		/// </summary>
		/// <param name="reference">optional object Reference</param>
		/// <param name="scroll">optional object Scroll</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Goto(object reference, object scroll)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reference, scroll);
			Invoker.Method(this, "Goto", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839232.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Goto()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Goto", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839232.aspx
		/// </summary>
		/// <param name="reference">optional object Reference</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Goto(object reference)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reference);
			Invoker.Method(this, "Goto", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840286.aspx
		/// </summary>
		/// <param name="helpFile">optional object HelpFile</param>
		/// <param name="helpContextID">optional object HelpContextID</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Help(object helpFile, object helpContextID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(helpFile, helpContextID);
			Invoker.Method(this, "Help", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840286.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Help()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Help", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840286.aspx
		/// </summary>
		/// <param name="helpFile">optional object HelpFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Help(object helpFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(helpFile);
			Invoker.Method(this, "Help", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823169.aspx
		/// </summary>
		/// <param name="inches">Double Inches</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Double InchesToPoints(Double inches)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(inches);
			object returnItem = Invoker.MethodReturn(this, "InchesToPoints", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839468.aspx
		/// </summary>
		/// <param name="prompt">string Prompt</param>
		/// <param name="title">optional object Title</param>
		/// <param name="_default">optional object Default</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="helpFile">optional object HelpFile</param>
		/// <param name="helpContextID">optional object HelpContextID</param>
		/// <param name="type">optional object Type</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object InputBox(string prompt, object title, object _default, object left, object top, object helpFile, object helpContextID, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(prompt, title, _default, left, top, helpFile, helpContextID, type);
			object returnItem = Invoker.MethodReturn(this, "InputBox", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839468.aspx
		/// </summary>
		/// <param name="prompt">string Prompt</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object InputBox(string prompt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(prompt);
			object returnItem = Invoker.MethodReturn(this, "InputBox", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839468.aspx
		/// </summary>
		/// <param name="prompt">string Prompt</param>
		/// <param name="title">optional object Title</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object InputBox(string prompt, object title)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(prompt, title);
			object returnItem = Invoker.MethodReturn(this, "InputBox", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839468.aspx
		/// </summary>
		/// <param name="prompt">string Prompt</param>
		/// <param name="title">optional object Title</param>
		/// <param name="_default">optional object Default</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object InputBox(string prompt, object title, object _default)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(prompt, title, _default);
			object returnItem = Invoker.MethodReturn(this, "InputBox", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839468.aspx
		/// </summary>
		/// <param name="prompt">string Prompt</param>
		/// <param name="title">optional object Title</param>
		/// <param name="_default">optional object Default</param>
		/// <param name="left">optional object Left</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object InputBox(string prompt, object title, object _default, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(prompt, title, _default, left);
			object returnItem = Invoker.MethodReturn(this, "InputBox", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839468.aspx
		/// </summary>
		/// <param name="prompt">string Prompt</param>
		/// <param name="title">optional object Title</param>
		/// <param name="_default">optional object Default</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object InputBox(string prompt, object title, object _default, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(prompt, title, _default, left, top);
			object returnItem = Invoker.MethodReturn(this, "InputBox", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839468.aspx
		/// </summary>
		/// <param name="prompt">string Prompt</param>
		/// <param name="title">optional object Title</param>
		/// <param name="_default">optional object Default</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="helpFile">optional object HelpFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object InputBox(string prompt, object title, object _default, object left, object top, object helpFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(prompt, title, _default, left, top, helpFile);
			object returnItem = Invoker.MethodReturn(this, "InputBox", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839468.aspx
		/// </summary>
		/// <param name="prompt">string Prompt</param>
		/// <param name="title">optional object Title</param>
		/// <param name="_default">optional object Default</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="helpFile">optional object HelpFile</param>
		/// <param name="helpContextID">optional object HelpContextID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object InputBox(string prompt, object title, object _default, object left, object top, object helpFile, object helpContextID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(prompt, title, _default, left, top, helpFile, helpContextID);
			object returnItem = Invoker.MethodReturn(this, "InputBox", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="statusBar">optional object StatusBar</param>
		/// <param name="helpContextID">optional object HelpContextID</param>
		/// <param name="helpFile">optional object HelpFile</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey, object shortcutKey, object category, object statusBar, object helpContextID, object helpFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey, shortcutKey, category, statusBar, helpContextID, helpFile);
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="statusBar">optional object StatusBar</param>
		/// <param name="helpContextID">optional object HelpContextID</param>
		/// <param name="helpFile">optional object HelpFile</param>
		/// <param name="argumentDescriptions">optional object ArgumentDescriptions</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey, object shortcutKey, object category, object statusBar, object helpContextID, object helpFile, object argumentDescriptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey, shortcutKey, category, statusBar, helpContextID, helpFile, argumentDescriptions);
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MacroOptions()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MacroOptions(object macro)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro);
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MacroOptions(object macro, object description)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description);
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MacroOptions(object macro, object description, object hasMenu)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu);
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MacroOptions(object macro, object description, object hasMenu, object menuText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText);
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey);
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey, object shortcutKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey, shortcutKey);
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey, object shortcutKey, object category)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey, shortcutKey, category);
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="statusBar">optional object StatusBar</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey, object shortcutKey, object category, object statusBar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey, shortcutKey, category, statusBar);
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838997.aspx
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="statusBar">optional object StatusBar</param>
		/// <param name="helpContextID">optional object HelpContextID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey, object shortcutKey, object category, object statusBar, object helpContextID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey, shortcutKey, category, statusBar, helpContextID);
			Invoker.Method(this, "MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820753.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MailLogoff()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MailLogoff", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193562.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="password">optional object Password</param>
		/// <param name="downloadNewMail">optional object DownloadNewMail</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MailLogon(object name, object password, object downloadNewMail)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, password, downloadNewMail);
			Invoker.Method(this, "MailLogon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193562.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MailLogon()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MailLogon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193562.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MailLogon(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "MailLogon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193562.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void MailLogon(object name, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, password);
			Invoker.Method(this, "MailLogon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192930.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook NextLetter()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NextLetter", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197461.aspx
		/// </summary>
		/// <param name="key">string Key</param>
		/// <param name="procedure">optional object Procedure</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OnKey(string key, object procedure)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key, procedure);
			Invoker.Method(this, "OnKey", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197461.aspx
		/// </summary>
		/// <param name="key">string Key</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OnKey(string key)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key);
			Invoker.Method(this, "OnKey", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834634.aspx
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="procedure">string Procedure</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OnRepeat(string text, string procedure)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, procedure);
			Invoker.Method(this, "OnRepeat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196165.aspx
		/// </summary>
		/// <param name="earliestTime">object EarliestTime</param>
		/// <param name="procedure">string Procedure</param>
		/// <param name="latestTime">optional object LatestTime</param>
		/// <param name="schedule">optional object Schedule</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OnTime(object earliestTime, string procedure, object latestTime, object schedule)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(earliestTime, procedure, latestTime, schedule);
			Invoker.Method(this, "OnTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196165.aspx
		/// </summary>
		/// <param name="earliestTime">object EarliestTime</param>
		/// <param name="procedure">string Procedure</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OnTime(object earliestTime, string procedure)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(earliestTime, procedure);
			Invoker.Method(this, "OnTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196165.aspx
		/// </summary>
		/// <param name="earliestTime">object EarliestTime</param>
		/// <param name="procedure">string Procedure</param>
		/// <param name="latestTime">optional object LatestTime</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OnTime(object earliestTime, string procedure, object latestTime)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(earliestTime, procedure, latestTime);
			Invoker.Method(this, "OnTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194135.aspx
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="procedure">string Procedure</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OnUndo(string text, string procedure)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, procedure);
			Invoker.Method(this, "OnUndo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839269.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Quit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835540.aspx
		/// </summary>
		/// <param name="basicCode">optional object BasicCode</param>
		/// <param name="xlmCode">optional object XlmCode</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void RecordMacro(object basicCode, object xlmCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(basicCode, xlmCode);
			Invoker.Method(this, "RecordMacro", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835540.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void RecordMacro()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RecordMacro", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835540.aspx
		/// </summary>
		/// <param name="basicCode">optional object BasicCode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void RecordMacro(object basicCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(basicCode);
			Invoker.Method(this, "RecordMacro", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837989.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool RegisterXLL(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			object returnItem = Invoker.MethodReturn(this, "RegisterXLL", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839236.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Repeat()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Repeat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void ResetTipWizard()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ResetTipWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Save(object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Save()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837602.aspx
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveWorkspace(object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "SaveWorkspace", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837602.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SaveWorkspace()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SaveWorkspace", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formatName">optional object FormatName</param>
		/// <param name="gallery">optional object Gallery</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SetDefaultChart(object formatName, object gallery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formatName, gallery);
			Invoker.Method(this, "SetDefaultChart", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SetDefaultChart()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SetDefaultChart", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formatName">optional object FormatName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void SetDefaultChart(object formatName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formatName);
			Invoker.Method(this, "SetDefaultChart", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838189.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Undo()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Undo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195441.aspx
		/// </summary>
		/// <param name="_volatile">optional object Volatile</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Volatile(object _volatile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_volatile);
			Invoker.Method(this, "Volatile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195441.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Volatile()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Volatile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="time">object Time</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _Wait(object time)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(time);
			Invoker.Method(this, "_Wait", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		/// <param name="arg30">optional object Arg30</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _WSFunction(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29);
			object returnItem = Invoker.MethodReturn(this, "_WSFunction", paramsArray);
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822851.aspx
		/// </summary>
		/// <param name="time">object Time</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Wait(object time)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(time);
			object returnItem = Invoker.MethodReturn(this, "Wait", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820792.aspx
		/// </summary>
		/// <param name="text">optional object Text</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string GetPhonetic(object text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			object returnItem = Invoker.MethodReturn(this, "GetPhonetic", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820792.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string GetPhonetic()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetPhonetic", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9)]
		public void Dummy12()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy12", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="p1">NetOffice.ExcelApi.PivotTable p1</param>
		/// <param name="p2">NetOffice.ExcelApi.PivotTable p2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void Dummy12(NetOffice.ExcelApi.PivotTable p1, NetOffice.ExcelApi.PivotTable p2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(p1, p2);
			Invoker.Method(this, "Dummy12", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194064.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void CalculateFull()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CalculateFull", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838801.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool FindFile()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FindFile", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		/// <param name="arg30">optional object Arg30</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Dummy13(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29);
			object returnItem = Invoker.MethodReturn(this, "Dummy13", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void Dummy14()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy14", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822609.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void CalculateFullRebuild()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CalculateFullRebuild", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840247.aspx
		/// </summary>
		/// <param name="keepAbort">optional object KeepAbort</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void CheckAbort(object keepAbort)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(keepAbort);
			Invoker.Method(this, "CheckAbort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840247.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void CheckAbort()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckAbort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194869.aspx
		/// </summary>
		/// <param name="xmlMap">optional object XmlMap</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public void DisplayXMLSourcePane(object xmlMap)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xmlMap);
			Invoker.Method(this, "DisplayXMLSourcePane", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194869.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public void DisplayXMLSourcePane()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DisplayXMLSourcePane", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="_object">object Object</param>
		/// <param name="iD">Int32 ID</param>
		/// <param name="arg">optional object arg</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public object Support(object _object, Int32 iD, object arg)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_object, iD, arg);
			object returnItem = Invoker.MethodReturn(this, "Support", paramsArray);
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
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="_object">object Object</param>
		/// <param name="iD">Int32 ID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public object Support(object _object, Int32 iD)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_object, iD);
			object returnItem = Invoker.MethodReturn(this, "Support", paramsArray);
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="grfCompareFunctions">Int32 grfCompareFunctions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object Dummy20(Int32 grfCompareFunctions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(grfCompareFunctions);
			object returnItem = Invoker.MethodReturn(this, "Dummy20", paramsArray);
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821008.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public void CalculateUntilAsyncQueriesDone()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CalculateUntilAsyncQueriesDone", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836804.aspx
		/// </summary>
		/// <param name="bstrUrl">string bstrUrl</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 SharePointVersion(string bstrUrl)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrUrl);
			object returnItem = Invoker.MethodReturn(this, "SharePointVersion", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="statusBar">optional object StatusBar</param>
		/// <param name="helpContextID">optional object HelpContextID</param>
		/// <param name="helpFile">optional object HelpFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void _MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey, object shortcutKey, object category, object statusBar, object helpContextID, object helpFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey, shortcutKey, category, statusBar, helpContextID, helpFile);
			Invoker.Method(this, "_MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void _MacroOptions()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void _MacroOptions(object macro)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro);
			Invoker.Method(this, "_MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void _MacroOptions(object macro, object description)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description);
			Invoker.Method(this, "_MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void _MacroOptions(object macro, object description, object hasMenu)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu);
			Invoker.Method(this, "_MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void _MacroOptions(object macro, object description, object hasMenu, object menuText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText);
			Invoker.Method(this, "_MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void _MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey);
			Invoker.Method(this, "_MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void _MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey, object shortcutKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey, shortcutKey);
			Invoker.Method(this, "_MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void _MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey, object shortcutKey, object category)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey, shortcutKey, category);
			Invoker.Method(this, "_MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="statusBar">optional object StatusBar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void _MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey, object shortcutKey, object category, object statusBar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey, shortcutKey, category, statusBar);
			Invoker.Method(this, "_MacroOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="macro">optional object Macro</param>
		/// <param name="description">optional object Description</param>
		/// <param name="hasMenu">optional object HasMenu</param>
		/// <param name="menuText">optional object MenuText</param>
		/// <param name="hasShortcutKey">optional object HasShortcutKey</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="category">optional object Category</param>
		/// <param name="statusBar">optional object StatusBar</param>
		/// <param name="helpContextID">optional object HelpContextID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public void _MacroOptions(object macro, object description, object hasMenu, object menuText, object hasShortcutKey, object shortcutKey, object category, object statusBar, object helpContextID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macro, description, hasMenu, menuText, hasShortcutKey, shortcutKey, category, statusBar, helpContextID);
			Invoker.Method(this, "_MacroOptions", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}