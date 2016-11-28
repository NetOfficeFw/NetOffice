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
	/// Interface IWorkbookEvents 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IWorkbookEvents : COMObject
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
                    _type = typeof(IWorkbookEvents);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IWorkbookEvents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookEvents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookEvents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookEvents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookEvents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookEvents() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookEvents(string progId) : base(progId)
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
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Open()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Activate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Activate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Deactivate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Deactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 BeforeClose(bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cancel);
			object returnItem = Invoker.MethodReturn(this, "BeforeClose", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="saveAsUI">bool SaveAsUI</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 BeforeSave(bool saveAsUI, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveAsUI, cancel);
			object returnItem = Invoker.MethodReturn(this, "BeforeSave", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 BeforePrint(bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cancel);
			object returnItem = Invoker.MethodReturn(this, "BeforePrint", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sh">object Sh</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 NewSheet(object sh)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh);
			object returnItem = Invoker.MethodReturn(this, "NewSheet", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 AddinInstall()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddinInstall", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 AddinUninstall()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddinUninstall", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wn">NetOffice.ExcelApi.Window Wn</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WindowResize(NetOffice.ExcelApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wn);
			object returnItem = Invoker.MethodReturn(this, "WindowResize", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wn">NetOffice.ExcelApi.Window Wn</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WindowActivate(NetOffice.ExcelApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wn);
			object returnItem = Invoker.MethodReturn(this, "WindowActivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wn">NetOffice.ExcelApi.Window Wn</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 WindowDeactivate(NetOffice.ExcelApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wn);
			object returnItem = Invoker.MethodReturn(this, "WindowDeactivate", paramsArray);
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
		/// <param name="target">NetOffice.ExcelApi.PivotTable Target</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 PivotTableCloseConnection(NetOffice.ExcelApi.PivotTable target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target);
			object returnItem = Invoker.MethodReturn(this, "PivotTableCloseConnection", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.PivotTable Target</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 PivotTableOpenConnection(NetOffice.ExcelApi.PivotTable target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target);
			object returnItem = Invoker.MethodReturn(this, "PivotTableOpenConnection", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="syncEventType">NetOffice.OfficeApi.Enums.MsoSyncEventType SyncEventType</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public Int32 Sync(NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(syncEventType);
			object returnItem = Invoker.MethodReturn(this, "Sync", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap Map</param>
		/// <param name="url">string Url</param>
		/// <param name="isRefresh">bool IsRefresh</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public Int32 BeforeXmlImport(NetOffice.ExcelApi.XmlMap map, string url, bool isRefresh, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(map, url, isRefresh, cancel);
			object returnItem = Invoker.MethodReturn(this, "BeforeXmlImport", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap Map</param>
		/// <param name="isRefresh">bool IsRefresh</param>
		/// <param name="result">NetOffice.ExcelApi.Enums.XlXmlImportResult Result</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public Int32 AfterXmlImport(NetOffice.ExcelApi.XmlMap map, bool isRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult result)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(map, isRefresh, result);
			object returnItem = Invoker.MethodReturn(this, "AfterXmlImport", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap Map</param>
		/// <param name="url">string Url</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public Int32 BeforeXmlExport(NetOffice.ExcelApi.XmlMap map, string url, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(map, url, cancel);
			object returnItem = Invoker.MethodReturn(this, "BeforeXmlExport", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap Map</param>
		/// <param name="url">string Url</param>
		/// <param name="result">NetOffice.ExcelApi.Enums.XlXmlExportResult Result</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public Int32 AfterXmlExport(NetOffice.ExcelApi.XmlMap map, string url, NetOffice.ExcelApi.Enums.XlXmlExportResult result)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(map, url, result);
			object returnItem = Invoker.MethodReturn(this, "AfterXmlExport", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="description">string Description</param>
		/// <param name="sheet">string Sheet</param>
		/// <param name="success">bool Success</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 RowsetComplete(string description, string sheet, bool success)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(description, sheet, success);
			object returnItem = Invoker.MethodReturn(this, "RowsetComplete", paramsArray);
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
		/// <param name="sh">object Sh</param>
		/// <param name="target">NetOffice.ExcelApi.PivotTable Target</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 SheetPivotTableChangeSync(object sh, NetOffice.ExcelApi.PivotTable target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sh, target);
			object returnItem = Invoker.MethodReturn(this, "SheetPivotTableChangeSync", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="success">bool Success</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 AfterSave(bool success)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(success);
			object returnItem = Invoker.MethodReturn(this, "AfterSave", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="ch">NetOffice.ExcelApi.Chart Ch</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 NewChart(NetOffice.ExcelApi.Chart ch)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(ch);
			object returnItem = Invoker.MethodReturn(this, "NewChart", paramsArray);
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
		/// <param name="changes">NetOffice.ExcelApi.ModelChanges Changes</param>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public Int32 ModelChange(NetOffice.ExcelApi.ModelChanges changes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(changes);
			object returnItem = Invoker.MethodReturn(this, "ModelChange", paramsArray);
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