using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// Interface IWorkbookEvents 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IWorkbookEvents : COMObject, NetOffice.ExcelApi.IWorkbookEvents
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.ExcelApi.IWorkbookEvents);
                return _contractType;
            }
        }
        private static Type _contractType;


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
                    _type = typeof(IWorkbookEvents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IWorkbookEvents() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 Open()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Open");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 Activate()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 Deactivate()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Deactivate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 BeforeClose(bool cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeforeClose", cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="saveAsUI">bool saveAsUI</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 BeforeSave(bool saveAsUI, bool cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeforeSave", saveAsUI, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 BeforePrint(bool cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeforePrint", cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 NewSheet(object sh)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "NewSheet", sh);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 AddinInstall()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AddinInstall");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 AddinUninstall()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AddinUninstall");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wn">NetOffice.ExcelApi.Window wn</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 WindowResize(NetOffice.ExcelApi.Window wn)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "WindowResize", wn);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wn">NetOffice.ExcelApi.Window wn</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 WindowActivate(NetOffice.ExcelApi.Window wn)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "WindowActivate", wn);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="wn">NetOffice.ExcelApi.Window wn</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 WindowDeactivate(NetOffice.ExcelApi.Window wn)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "WindowDeactivate", wn);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 SheetSelectionChange(object sh, NetOffice.ExcelApi.Range target)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetSelectionChange", sh, target);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 SheetBeforeDoubleClick(object sh, NetOffice.ExcelApi.Range target, bool cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetBeforeDoubleClick", sh, target, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 SheetBeforeRightClick(object sh, NetOffice.ExcelApi.Range target, bool cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetBeforeRightClick", sh, target, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 SheetActivate(object sh)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetActivate", sh);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 SheetDeactivate(object sh)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetDeactivate", sh);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 SheetCalculate(object sh)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetCalculate", sh);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 SheetChange(object sh, NetOffice.ExcelApi.Range target)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetChange", sh, target);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.Hyperlink target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 SheetFollowHyperlink(object sh, NetOffice.ExcelApi.Hyperlink target)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetFollowHyperlink", sh, target);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual Int32 SheetPivotTableUpdate(object sh, NetOffice.ExcelApi.PivotTable target)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetPivotTableUpdate", sh, target);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual Int32 PivotTableCloseConnection(NetOffice.ExcelApi.PivotTable target)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableCloseConnection", target);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual Int32 PivotTableOpenConnection(NetOffice.ExcelApi.PivotTable target)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PivotTableOpenConnection", target);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="syncEventType">NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual Int32 Sync(NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Sync", syncEventType);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="url">string url</param>
		/// <param name="isRefresh">bool isRefresh</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual Int32 BeforeXmlImport(NetOffice.ExcelApi.XmlMap map, string url, bool isRefresh, bool cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeforeXmlImport", map, url, isRefresh, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="isRefresh">bool isRefresh</param>
		/// <param name="result">NetOffice.ExcelApi.Enums.XlXmlImportResult result</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual Int32 AfterXmlImport(NetOffice.ExcelApi.XmlMap map, bool isRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult result)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AfterXmlImport", map, isRefresh, result);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="url">string url</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual Int32 BeforeXmlExport(NetOffice.ExcelApi.XmlMap map, string url, bool cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeforeXmlExport", map, url, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="url">string url</param>
		/// <param name="result">NetOffice.ExcelApi.Enums.XlXmlExportResult result</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual Int32 AfterXmlExport(NetOffice.ExcelApi.XmlMap map, string url, NetOffice.ExcelApi.Enums.XlXmlExportResult result)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AfterXmlExport", map, url, result);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="description">string description</param>
		/// <param name="sheet">string sheet</param>
		/// <param name="success">bool success</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual Int32 RowsetComplete(string description, string sheet, bool success)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RowsetComplete", description, sheet, success);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="targetRange">NetOffice.ExcelApi.Range targetRange</param>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual Int32 SheetPivotTableAfterValueChange(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetPivotTableAfterValueChange", sh, targetPivotTable, targetRange);
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
		public virtual Int32 SheetPivotTableBeforeAllocateChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetPivotTableBeforeAllocateChanges", new object[]{ sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel });
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
		public virtual Int32 SheetPivotTableBeforeCommitChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetPivotTableBeforeCommitChanges", new object[]{ sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="valueChangeStart">Int32 valueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual Int32 SheetPivotTableBeforeDiscardChanges(object sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetPivotTableBeforeDiscardChanges", sh, targetPivotTable, valueChangeStart, valueChangeEnd);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual Int32 SheetPivotTableChangeSync(object sh, NetOffice.ExcelApi.PivotTable target)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetPivotTableChangeSync", sh, target);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="success">bool success</param>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual Int32 AfterSave(bool success)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AfterSave", success);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="ch">NetOffice.ExcelApi.Chart ch</param>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual Int32 NewChart(NetOffice.ExcelApi.Chart ch)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "NewChart", ch);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 15, 16)]
		public virtual Int32 SheetLensGalleryRenderComplete(object sh)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetLensGalleryRenderComplete", sh);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="sh">object sh</param>
		/// <param name="target">NetOffice.ExcelApi.TableObject target</param>
		[SupportByVersion("Excel", 15, 16)]
		public virtual Int32 SheetTableUpdate(object sh, NetOffice.ExcelApi.TableObject target)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetTableUpdate", sh, target);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="changes">NetOffice.ExcelApi.ModelChanges changes</param>
		[SupportByVersion("Excel", 15, 16)]
		public virtual Int32 ModelChange(NetOffice.ExcelApi.ModelChanges changes)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ModelChange", changes);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="sh">object sh</param>
		[SupportByVersion("Excel", 15, 16)]
		public virtual Int32 SheetBeforeDelete(object sh)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SheetBeforeDelete", sh);
		}

		#endregion

		#pragma warning restore
	}
}

