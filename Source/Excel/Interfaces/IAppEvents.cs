using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.ExcelApi
{
	///<summary>
	/// Interface IAppEvents 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IAppEvents : COMObject
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
                    _type = typeof(IAppEvents);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IAppEvents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAppEvents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAppEvents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAppEvents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAppEvents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAppEvents() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAppEvents(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 NewWorkbook(NetOffice.ExcelApi.Workbook wb)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb);
			object returnItem = Invoker.MethodReturn(this, "NewWorkbook", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range Target</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetSelectionChange(object sh, NetOffice.ExcelApi.Range target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, target);
			object returnItem = Invoker.MethodReturn(this, "SheetSelectionChange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range Target</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetBeforeDoubleClick(object sh, NetOffice.ExcelApi.Range target, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, target, cancel);
			object returnItem = Invoker.MethodReturn(this, "SheetBeforeDoubleClick", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range Target</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetBeforeRightClick(object sh, NetOffice.ExcelApi.Range target, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, target, cancel);
			object returnItem = Invoker.MethodReturn(this, "SheetBeforeRightClick", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetActivate(object sh)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh);
			object returnItem = Invoker.MethodReturn(this, "SheetActivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetDeactivate(object sh)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh);
			object returnItem = Invoker.MethodReturn(this, "SheetDeactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetCalculate(object sh)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh);
			object returnItem = Invoker.MethodReturn(this, "SheetCalculate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range Target</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetChange(object sh, NetOffice.ExcelApi.Range target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, target);
			object returnItem = Invoker.MethodReturn(this, "SheetChange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookOpen(NetOffice.ExcelApi.Workbook wb)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb);
			object returnItem = Invoker.MethodReturn(this, "WorkbookOpen", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookActivate(NetOffice.ExcelApi.Workbook wb)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb);
			object returnItem = Invoker.MethodReturn(this, "WorkbookActivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookDeactivate(NetOffice.ExcelApi.Workbook wb)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb);
			object returnItem = Invoker.MethodReturn(this, "WorkbookDeactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookBeforeClose(NetOffice.ExcelApi.Workbook wb, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, cancel);
			object returnItem = Invoker.MethodReturn(this, "WorkbookBeforeClose", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="saveAsUI">bool SaveAsUI</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookBeforeSave(NetOffice.ExcelApi.Workbook wb, bool saveAsUI, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, saveAsUI, cancel);
			object returnItem = Invoker.MethodReturn(this, "WorkbookBeforeSave", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookBeforePrint(NetOffice.ExcelApi.Workbook wb, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, cancel);
			object returnItem = Invoker.MethodReturn(this, "WorkbookBeforePrint", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="sh">object Sh</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookNewSheet(NetOffice.ExcelApi.Workbook wb, object sh)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, sh);
			object returnItem = Invoker.MethodReturn(this, "WorkbookNewSheet", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookAddinInstall(NetOffice.ExcelApi.Workbook wb)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb);
			object returnItem = Invoker.MethodReturn(this, "WorkbookAddinInstall", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookAddinUninstall(NetOffice.ExcelApi.Workbook wb)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb);
			object returnItem = Invoker.MethodReturn(this, "WorkbookAddinUninstall", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="wn">NetOffice.ExcelApi.Window Wn</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WindowResize(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, wn);
			object returnItem = Invoker.MethodReturn(this, "WindowResize", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="wn">NetOffice.ExcelApi.Window Wn</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WindowActivate(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, wn);
			object returnItem = Invoker.MethodReturn(this, "WindowActivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="wn">NetOffice.ExcelApi.Window Wn</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WindowDeactivate(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, wn);
			object returnItem = Invoker.MethodReturn(this, "WindowDeactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		/// <param name="target">NetOffice.ExcelApi.Hyperlink Target</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetFollowHyperlink(object sh, NetOffice.ExcelApi.Hyperlink target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, target);
			object returnItem = Invoker.MethodReturn(this, "SheetFollowHyperlink", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		/// <param name="target">NetOffice.ExcelApi.PivotTable Target</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 SheetPivotTableUpdate(object sh, NetOffice.ExcelApi.PivotTable target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, target);
			object returnItem = Invoker.MethodReturn(this, "SheetPivotTableUpdate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="target">NetOffice.ExcelApi.PivotTable Target</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 WorkbookPivotTableCloseConnection(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, target);
			object returnItem = Invoker.MethodReturn(this, "WorkbookPivotTableCloseConnection", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="target">NetOffice.ExcelApi.PivotTable Target</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 WorkbookPivotTableOpenConnection(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, target);
			object returnItem = Invoker.MethodReturn(this, "WorkbookPivotTableOpenConnection", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="syncEventType">NetOffice.OfficeApi.Enums.MsoSyncEventType SyncEventType</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public Int32 WorkbookSync(NetOffice.ExcelApi.Workbook wb, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, syncEventType);
			object returnItem = Invoker.MethodReturn(this, "WorkbookSync", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="map">NetOffice.ExcelApi.XmlMap Map</param>
		/// <param name="url">string Url</param>
		/// <param name="isRefresh">bool IsRefresh</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public Int32 WorkbookBeforeXmlImport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, bool isRefresh, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, map, url, isRefresh, cancel);
			object returnItem = Invoker.MethodReturn(this, "WorkbookBeforeXmlImport", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="map">NetOffice.ExcelApi.XmlMap Map</param>
		/// <param name="isRefresh">bool IsRefresh</param>
		/// <param name="result">NetOffice.ExcelApi.Enums.XlXmlImportResult Result</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public Int32 WorkbookAfterXmlImport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, bool isRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult result)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, map, isRefresh, result);
			object returnItem = Invoker.MethodReturn(this, "WorkbookAfterXmlImport", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="map">NetOffice.ExcelApi.XmlMap Map</param>
		/// <param name="url">string Url</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public Int32 WorkbookBeforeXmlExport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, map, url, cancel);
			object returnItem = Invoker.MethodReturn(this, "WorkbookBeforeXmlExport", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="map">NetOffice.ExcelApi.XmlMap Map</param>
		/// <param name="url">string Url</param>
		/// <param name="result">NetOffice.ExcelApi.Enums.XlXmlExportResult Result</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public Int32 WorkbookAfterXmlExport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, NetOffice.ExcelApi.Enums.XlXmlExportResult result)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, map, url, result);
			object returnItem = Invoker.MethodReturn(this, "WorkbookAfterXmlExport", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="description">string Description</param>
		/// <param name="sheet">string Sheet</param>
		/// <param name="success">bool Success</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 WorkbookRowsetComplete(NetOffice.ExcelApi.Workbook wb, string description, string sheet, bool success)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, description, sheet, success);
			object returnItem = Invoker.MethodReturn(this, "WorkbookRowsetComplete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 AfterCalculate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AfterCalculate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable TargetPivotTable</param>
		/// <param name="targetRange">NetOffice.ExcelApi.Range TargetRange</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 SheetPivotTableAfterValueChange(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, targetPivotTable, targetRange);
			object returnItem = Invoker.MethodReturn(this, "SheetPivotTableAfterValueChange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable TargetPivotTable</param>
		/// <param name="valueChangeStart">Int32 ValueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 ValueChangeEnd</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 SheetPivotTableBeforeAllocateChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
			object returnItem = Invoker.MethodReturn(this, "SheetPivotTableBeforeAllocateChanges", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable TargetPivotTable</param>
		/// <param name="valueChangeStart">Int32 ValueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 ValueChangeEnd</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 SheetPivotTableBeforeCommitChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
			object returnItem = Invoker.MethodReturn(this, "SheetPivotTableBeforeCommitChanges", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable TargetPivotTable</param>
		/// <param name="valueChangeStart">Int32 ValueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 ValueChangeEnd</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 SheetPivotTableBeforeDiscardChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd);
			object returnItem = Invoker.MethodReturn(this, "SheetPivotTableBeforeDiscardChanges", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow Pvw</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowOpen(NetOffice.ExcelApi.ProtectedViewWindow pvw)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvw);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowOpen", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow Pvw</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowBeforeEdit(NetOffice.ExcelApi.ProtectedViewWindow pvw, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvw, cancel);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowBeforeEdit", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow Pvw</param>
		/// <param name="reason">NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason Reason</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowBeforeClose(NetOffice.ExcelApi.ProtectedViewWindow pvw, NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason reason, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvw, reason, cancel);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowBeforeClose", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow Pvw</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowResize(NetOffice.ExcelApi.ProtectedViewWindow pvw)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvw);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowResize", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow Pvw</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowActivate(NetOffice.ExcelApi.ProtectedViewWindow pvw)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvw);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowActivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow Pvw</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowDeactivate(NetOffice.ExcelApi.ProtectedViewWindow pvw)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvw);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowDeactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="success">bool Success</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 WorkbookAfterSave(NetOffice.ExcelApi.Workbook wb, bool success)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, success);
			object returnItem = Invoker.MethodReturn(this, "WorkbookAfterSave", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="ch">NetOffice.ExcelApi.Chart Ch</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 WorkbookNewChart(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Chart ch)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, ch);
			object returnItem = Invoker.MethodReturn(this, "WorkbookNewChart", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public Int32 SheetLensGalleryRenderComplete(object sh)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh);
			object returnItem = Invoker.MethodReturn(this, "SheetLensGalleryRenderComplete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		/// <param name="target">NetOffice.ExcelApi.TableObject Target</param>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public Int32 SheetTableUpdate(object sh, NetOffice.ExcelApi.TableObject target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, target);
			object returnItem = Invoker.MethodReturn(this, "SheetTableUpdate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// 
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook Wb</param>
		/// <param name="changes">NetOffice.ExcelApi.ModelChanges Changes</param>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public Int32 WorkbookModelChange(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.ModelChanges changes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wb, changes);
			object returnItem = Invoker.MethodReturn(this, "WorkbookModelChange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public Int32 SheetBeforeDelete(object sh)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh);
			object returnItem = Invoker.MethodReturn(this, "SheetBeforeDelete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}