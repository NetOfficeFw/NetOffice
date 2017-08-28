using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IAppEvents 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IAppEvents : COMObject
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
                    _type = typeof(IAppEvents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IAppEvents(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 NewWorkbook(NetOffice.ExcelApi.Workbook wb)
		{
			return Factory.ExecuteInt32MethodGet(this, "NewWorkbook", wb);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetSelectionChange(object sh, NetOffice.ExcelApi.Range target)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetSelectionChange", sh, target);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetBeforeDoubleClick(object sh, NetOffice.ExcelApi.Range target, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetBeforeDoubleClick", sh, target, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetBeforeRightClick(object sh, NetOffice.ExcelApi.Range target, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetBeforeRightClick", sh, target, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetActivate(object sh)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetActivate", sh);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetDeactivate(object sh)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetDeactivate", sh);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetCalculate(object sh)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetCalculate", sh);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetChange(object sh, NetOffice.ExcelApi.Range target)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetChange", sh, target);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookOpen(NetOffice.ExcelApi.Workbook wb)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookOpen", wb);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookActivate(NetOffice.ExcelApi.Workbook wb)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookActivate", wb);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookDeactivate(NetOffice.ExcelApi.Workbook wb)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookDeactivate", wb);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookBeforeClose(NetOffice.ExcelApi.Workbook wb, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookBeforeClose", wb, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="saveAsUI">bool saveAsUI</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookBeforeSave(NetOffice.ExcelApi.Workbook wb, bool saveAsUI, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookBeforeSave", wb, saveAsUI, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookBeforePrint(NetOffice.ExcelApi.Workbook wb, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookBeforePrint", wb, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookNewSheet(NetOffice.ExcelApi.Workbook wb, object sh)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookNewSheet", wb, sh);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookAddinInstall(NetOffice.ExcelApi.Workbook wb)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookAddinInstall", wb);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WorkbookAddinUninstall(NetOffice.ExcelApi.Workbook wb)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookAddinUninstall", wb);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="wn">NetOffice.ExcelApi.Window wn</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WindowResize(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn)
		{
			return Factory.ExecuteInt32MethodGet(this, "WindowResize", wb, wn);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="wn">NetOffice.ExcelApi.Window wn</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WindowActivate(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn)
		{
			return Factory.ExecuteInt32MethodGet(this, "WindowActivate", wb, wn);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="wn">NetOffice.ExcelApi.Window wn</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 WindowDeactivate(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn)
		{
			return Factory.ExecuteInt32MethodGet(this, "WindowDeactivate", wb, wn);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Hyperlink target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SheetFollowHyperlink(object sh, NetOffice.ExcelApi.Hyperlink target)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetFollowHyperlink", sh, target);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 SheetPivotTableUpdate(object sh, NetOffice.ExcelApi.PivotTable target)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetPivotTableUpdate", sh, target);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 WorkbookPivotTableCloseConnection(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookPivotTableCloseConnection", wb, target);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 WorkbookPivotTableOpenConnection(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookPivotTableOpenConnection", wb, target);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="syncEventType">NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 WorkbookSync(NetOffice.ExcelApi.Workbook wb, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookSync", wb, syncEventType);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="url">string url</param>
		/// <param name="isRefresh">bool isRefresh</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 WorkbookBeforeXmlImport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, bool isRefresh, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookBeforeXmlImport", new object[]{ wb, map, url, isRefresh, cancel });
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="isRefresh">bool isRefresh</param>
		/// <param name="result">NetOffice.ExcelApi.Enums.XlXmlImportResult result</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 WorkbookAfterXmlImport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, bool isRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult result)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookAfterXmlImport", wb, map, isRefresh, result);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="url">string url</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 WorkbookBeforeXmlExport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookBeforeXmlExport", wb, map, url, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="url">string url</param>
		/// <param name="result">NetOffice.ExcelApi.Enums.XlXmlExportResult result</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 WorkbookAfterXmlExport(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, NetOffice.ExcelApi.Enums.XlXmlExportResult result)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookAfterXmlExport", wb, map, url, result);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="description">string description</param>
		/// <param name="sheet">string sheet</param>
		/// <param name="success">bool success</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 WorkbookRowsetComplete(NetOffice.ExcelApi.Workbook wb, string description, string sheet, bool success)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookRowsetComplete", wb, description, sheet, success);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 AfterCalculate()
		{
			return Factory.ExecuteInt32MethodGet(this, "AfterCalculate");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="targetRange">NetOffice.ExcelApi.Range targetRange</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 SheetPivotTableAfterValueChange(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetPivotTableAfterValueChange", sh, targetPivotTable, targetRange);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="valueChangeStart">Int32 valueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 SheetPivotTableBeforeAllocateChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetPivotTableBeforeAllocateChanges", new object[]{ sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="valueChangeStart">Int32 valueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 SheetPivotTableBeforeCommitChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetPivotTableBeforeCommitChanges", new object[]{ sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="valueChangeStart">Int32 valueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 SheetPivotTableBeforeDiscardChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetPivotTableBeforeDiscardChanges", sh, targetPivotTable, valueChangeStart, valueChangeEnd);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowOpen(NetOffice.ExcelApi.ProtectedViewWindow pvw)
		{
			return Factory.ExecuteInt32MethodGet(this, "ProtectedViewWindowOpen", pvw);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowBeforeEdit(NetOffice.ExcelApi.ProtectedViewWindow pvw, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "ProtectedViewWindowBeforeEdit", pvw, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
		/// <param name="reason">NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason reason</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowBeforeClose(NetOffice.ExcelApi.ProtectedViewWindow pvw, NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason reason, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "ProtectedViewWindowBeforeClose", pvw, reason, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowResize(NetOffice.ExcelApi.ProtectedViewWindow pvw)
		{
			return Factory.ExecuteInt32MethodGet(this, "ProtectedViewWindowResize", pvw);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowActivate(NetOffice.ExcelApi.ProtectedViewWindow pvw)
		{
			return Factory.ExecuteInt32MethodGet(this, "ProtectedViewWindowActivate", pvw);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="pvw">NetOffice.ExcelApi.ProtectedViewWindow pvw</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 ProtectedViewWindowDeactivate(NetOffice.ExcelApi.ProtectedViewWindow pvw)
		{
			return Factory.ExecuteInt32MethodGet(this, "ProtectedViewWindowDeactivate", pvw);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="success">bool success</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 WorkbookAfterSave(NetOffice.ExcelApi.Workbook wb, bool success)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookAfterSave", wb, success);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="ch">NetOffice.ExcelApi.Chart ch</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 WorkbookNewChart(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Chart ch)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookNewChart", wb, ch);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 15, 16)]
		public Int32 SheetLensGalleryRenderComplete(object sh)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetLensGalleryRenderComplete", sh);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.TableObject target</param>
		[SupportByVersion("Excel", 15, 16)]
		public Int32 SheetTableUpdate(object sh, NetOffice.ExcelApi.TableObject target)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetTableUpdate", sh, target);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="wb">NetOffice.ExcelApi.Workbook wb</param>
		/// <param name="changes">NetOffice.ExcelApi.ModelChanges changes</param>
		[SupportByVersion("Excel", 15, 16)]
		public Int32 WorkbookModelChange(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.ModelChanges changes)
		{
			return Factory.ExecuteInt32MethodGet(this, "WorkbookModelChange", wb, changes);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 15, 16)]
		public Int32 SheetBeforeDelete(object sh)
		{
			return Factory.ExecuteInt32MethodGet(this, "SheetBeforeDelete", sh);
		}

		#endregion

		#pragma warning restore
	}
}
