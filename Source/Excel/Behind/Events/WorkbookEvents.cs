using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.Behind.EventContracts
{
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class WorkbookEvents_SinkHelper : SinkHelper, NetOffice.ExcelApi.EventContracts.WorkbookEvents
    {
        #region Static

        public static readonly string Id = "00024412-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public WorkbookEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region WorkbookEvents

        public void Open()
        {
            if (!Validate("Open"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Open", ref paramsArray);
        }

        public void Activate()
        {
            if (!Validate("Activate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Activate", ref paramsArray);
        }

        public void Deactivate()
        {
            if (!Validate("Deactivate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Deactivate", ref paramsArray);
        }

        public void BeforeClose([In] [Out] ref object cancel)
        {
            if (!Validate("BeforeClose"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
            paramsArray.SetValue(cancel, 0);
            EventBinding.RaiseCustomEvent("BeforeClose", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        public void BeforeSave([In] object saveAsUI, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeSave"))
            {
                Invoker.ReleaseParamsArray(saveAsUI, cancel);
                return;
            }

            bool newSaveAsUI = ToBoolean(saveAsUI);
            object[] paramsArray = new object[2];
            paramsArray[0] = newSaveAsUI;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("BeforeSave", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void BeforePrint([In] [Out] ref object cancel)
        {
            if (!Validate("BeforePrint"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
            paramsArray.SetValue(cancel, 0);
            EventBinding.RaiseCustomEvent("BeforePrint", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        public void NewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("NewSheet"))
            {
                Invoker.ReleaseParamsArray(sh);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSh;
            EventBinding.RaiseCustomEvent("NewSheet", ref paramsArray);
        }

        public void AddinInstall()
        {
            if (!Validate("AddinInstall"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("AddinInstall", ref paramsArray);
        }

        public void AddinUninstall()
        {
            if (!Validate("AddinUninstall"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("AddinUninstall", ref paramsArray);
        }

        public void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowResize"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.ExcelApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Window>(EventClass, wn, typeof(NetOffice.ExcelApi.Window));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWn;
            EventBinding.RaiseCustomEvent("WindowResize", ref paramsArray);
        }

        public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowActivate"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.ExcelApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Window>(EventClass, wn, typeof(NetOffice.ExcelApi.Window));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWn;
            EventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
        }

        public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.ExcelApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Window>(EventClass, wn, typeof(NetOffice.ExcelApi.Window));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWn;
            EventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
        }

        public void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetSelectionChange"))
            {
                Invoker.ReleaseParamsArray(sh, target);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, typeof(NetOffice.ExcelApi.Range));
            object[] paramsArray = new object[2];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("SheetSelectionChange", ref paramsArray);
        }

        public void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
        {
            if (!Validate("SheetBeforeDoubleClick"))
            {
                Invoker.ReleaseParamsArray(sh, target, cancel);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, typeof(NetOffice.ExcelApi.Range));
            object[] paramsArray = new object[3];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("SheetBeforeDoubleClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        public void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
        {
            if (!Validate("SheetBeforeRightClick"))
            {
                Invoker.ReleaseParamsArray(sh, target, cancel);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, typeof(NetOffice.ExcelApi.Range));
            object[] paramsArray = new object[3];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("SheetBeforeRightClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        public void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("SheetActivate"))
            {
                Invoker.ReleaseParamsArray(sh);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSh;
            EventBinding.RaiseCustomEvent("SheetActivate", ref paramsArray);
        }

        public void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("SheetDeactivate"))
            {
                Invoker.ReleaseParamsArray(sh);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSh;
            EventBinding.RaiseCustomEvent("SheetDeactivate", ref paramsArray);
        }

        public void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("SheetCalculate"))
            {
                Invoker.ReleaseParamsArray(sh);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSh;
            EventBinding.RaiseCustomEvent("SheetCalculate", ref paramsArray);
        }

        public void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetCalculate"))
            {
                Invoker.ReleaseParamsArray(sh, target);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, typeof(NetOffice.ExcelApi.Range));
            object[] paramsArray = new object[2];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("SheetChange", ref paramsArray);
        }

        public void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetFollowHyperlink"))
            {
                Invoker.ReleaseParamsArray(sh, target);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.Hyperlink newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Hyperlink>(EventClass, target, typeof(NetOffice.ExcelApi.Hyperlink));
            object[] paramsArray = new object[2];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("SheetFollowHyperlink", ref paramsArray);
        }

        public void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetPivotTableUpdate"))
            {
                Invoker.ReleaseParamsArray(sh, target);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, typeof(NetOffice.ExcelApi.PivotTable));
            object[] paramsArray = new object[2];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("SheetPivotTableUpdate", ref paramsArray);
        }

        public void PivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("PivotTableCloseConnection"))
            {
                Invoker.ReleaseParamsArray(target);
                return;
            }

            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, typeof(NetOffice.ExcelApi.PivotTable));
            object[] paramsArray = new object[1];
            paramsArray[0] = newTarget;
            EventBinding.RaiseCustomEvent("PivotTableCloseConnection", ref paramsArray);
        }

        public void PivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("PivotTableOpenConnection"))
            {
                Invoker.ReleaseParamsArray(target);
                return;
            }

            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, typeof(NetOffice.ExcelApi.PivotTable));
            object[] paramsArray = new object[1];
            paramsArray[0] = newTarget;
            EventBinding.RaiseCustomEvent("PivotTableOpenConnection", ref paramsArray);
        }

        public void Sync([In] object syncEventType)
        {
            if (!Validate("Sync"))
            {
                Invoker.ReleaseParamsArray(syncEventType);
                return;
            }

            NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSyncEventType;
            EventBinding.RaiseCustomEvent("Sync", ref paramsArray);
        }

        public void BeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeXmlImports"))
            {
                Invoker.ReleaseParamsArray(map, url, isRefresh, cancel);
                return;
            }

            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, typeof(NetOffice.ExcelApi.XmlMap));
            string newUrl = ToString(url);
            bool newIsRefresh = ToBoolean(isRefresh);
            object[] paramsArray = new object[4];
            paramsArray[0] = newMap;
            paramsArray[1] = newUrl;
            paramsArray[2] = newIsRefresh;
            paramsArray.SetValue(cancel, 3);
            EventBinding.RaiseCustomEvent("BeforeXmlImport", ref paramsArray);

            cancel = ToBoolean(paramsArray[3]);
        }

        public void AfterXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object isRefresh, [In] object result)
        {
            if (!Validate("AfterXmlImport"))
            {
                Invoker.ReleaseParamsArray(map, isRefresh, result);
                return;
            }

            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, typeof(NetOffice.ExcelApi.XmlMap));
            bool newIsRefresh = ToBoolean(isRefresh);
            NetOffice.ExcelApi.Enums.XlXmlImportResult newResult = (NetOffice.ExcelApi.Enums.XlXmlImportResult)result;
            object[] paramsArray = new object[3];
            paramsArray[0] = newMap;
            paramsArray[1] = newIsRefresh;
            paramsArray[2] = newResult;
            EventBinding.RaiseCustomEvent("AfterXmlImport", ref paramsArray);
        }

        public void BeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeXmlExport"))
            {
                Invoker.ReleaseParamsArray(map, url, cancel);
                return;
            }

            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, typeof(NetOffice.ExcelApi.XmlMap));
            string newUrl = ToString(url);
            object[] paramsArray = new object[3];
            paramsArray[0] = newMap;
            paramsArray[1] = newUrl;
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("BeforeXmlExport", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        public void AfterXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object result)
        {
            if (!Validate("AfterXmlExport"))
            {
                Invoker.ReleaseParamsArray(map, url, result);
                return;
            }

            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, typeof(NetOffice.ExcelApi.XmlMap));
            string newUrl = ToString(url);
            NetOffice.ExcelApi.Enums.XlXmlExportResult newResult = (NetOffice.ExcelApi.Enums.XlXmlExportResult)result;
            object[] paramsArray = new object[3];
            paramsArray[0] = newMap;
            paramsArray[1] = newUrl;
            paramsArray[2] = newResult;
            EventBinding.RaiseCustomEvent("AfterXmlExport", ref paramsArray);
        }

        public void RowsetComplete([In] object description, [In] object sheet, [In] object success)
        {
            if (!Validate("RowsetComplete"))
            {
                Invoker.ReleaseParamsArray(description, sheet, success);
                return;
            }

            string newDescription = ToString(description);
            string newSheet = ToString(sheet);
            bool newSuccess = ToBoolean(success);
            object[] paramsArray = new object[3];
            paramsArray[0] = newDescription;
            paramsArray[1] = newSheet;
            paramsArray[2] = newSuccess;
            EventBinding.RaiseCustomEvent("RowsetComplete", ref paramsArray);
        }

        public void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange)
        {
            if (!Validate("SheetPivotTableAfterValueChange"))
            {
                Invoker.ReleaseParamsArray(sh, targetPivotTable, targetRange);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, typeof(NetOffice.ExcelApi.PivotTable));
            NetOffice.ExcelApi.Range newTargetRange = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, targetRange, typeof(NetOffice.ExcelApi.Range));
            object[] paramsArray = new object[3];
            paramsArray[0] = newSh;
            paramsArray[1] = newTargetPivotTable;
            paramsArray[2] = newTargetRange;
            EventBinding.RaiseCustomEvent("SheetPivotTableAfterValueChange", ref paramsArray);
        }

        public void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
        {
            if (!Validate("SheetPivotTableBeforeAllocateChanges"))
            {
                Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, typeof(NetOffice.ExcelApi.PivotTable));
            Int32 newValueChangeStart = ToInt32(valueChangeStart);
            Int32 newValueChangeEnd = ToInt32(valueChangeEnd);
            object[] paramsArray = new object[5];
            paramsArray[0] = newSh;
            paramsArray[1] = newTargetPivotTable;
            paramsArray[2] = newValueChangeStart;
            paramsArray[3] = newValueChangeEnd;
            paramsArray.SetValue(cancel, 4);
            EventBinding.RaiseCustomEvent("SheetPivotTableBeforeAllocateChanges", ref paramsArray);

            cancel = ToBoolean(paramsArray[4]);
        }

        public void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
        {
            if (!Validate("SheetPivotTableBeforeAllocateChanges"))
            {
                Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, typeof(NetOffice.ExcelApi.PivotTable));
            Int32 newValueChangeStart = ToInt32(valueChangeStart);
            Int32 newValueChangeEnd = ToInt32(valueChangeEnd);
            object[] paramsArray = new object[5];
            paramsArray[0] = newSh;
            paramsArray[1] = newTargetPivotTable;
            paramsArray[2] = newValueChangeStart;
            paramsArray[3] = newValueChangeEnd;
            paramsArray.SetValue(cancel, 4);
            EventBinding.RaiseCustomEvent("SheetPivotTableBeforeCommitChanges", ref paramsArray);

            cancel = ToBoolean(paramsArray[4]);
        }

        public void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd)
        {
            if (!Validate("SheetPivotTableBeforeAllocateChanges"))
            {
                Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, typeof(NetOffice.ExcelApi.PivotTable));
            Int32 newValueChangeStart = ToInt32(valueChangeStart);
            Int32 newValueChangeEnd = ToInt32(valueChangeEnd);
            object[] paramsArray = new object[4];
            paramsArray[0] = newSh;
            paramsArray[1] = newTargetPivotTable;
            paramsArray[2] = newValueChangeStart;
            paramsArray[3] = newValueChangeEnd;
            EventBinding.RaiseCustomEvent("SheetPivotTableBeforeDiscardChanges", ref paramsArray);
        }

        public void SheetPivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetPivotTableBeforeAllocateChanges"))
            {
                Invoker.ReleaseParamsArray(sh, target);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, typeof(NetOffice.ExcelApi.PivotTable));
            object[] paramsArray = new object[2];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("SheetPivotTableChangeSync", ref paramsArray);
        }

        public void AfterSave([In] object success)
        {
            if (!Validate("AfterSave"))
            {
                Invoker.ReleaseParamsArray(success);
                return;
            }

            bool newSuccess = ToBoolean(success);
            object[] paramsArray = new object[1];
            paramsArray[0] = newSuccess;
            EventBinding.RaiseCustomEvent("AfterSave", ref paramsArray);
        }

        public void NewChart([In, MarshalAs(UnmanagedType.IDispatch)] object ch)
        {
            if (!Validate("NewChart"))
            {
                Invoker.ReleaseParamsArray(ch);
                return;
            }

            NetOffice.ExcelApi.Chart newCh = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Chart>(EventClass, ch, typeof(NetOffice.ExcelApi.Chart));
            object[] paramsArray = new object[1];
            paramsArray[0] = newCh;
            EventBinding.RaiseCustomEvent("NewChart", ref paramsArray);
        }

        public void SheetLensGalleryRenderComplete([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("SheetLensGalleryRenderComplete"))
            {
                Invoker.ReleaseParamsArray(sh);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSh;
            EventBinding.RaiseCustomEvent("SheetLensGalleryRenderComplete", ref paramsArray);
        }

        public void SheetTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetTableUpdate"))
            {
                Invoker.ReleaseParamsArray(sh, target);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.TableObject newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.TableObject>(EventClass, target, typeof(NetOffice.ExcelApi.TableObject));
            object[] paramsArray = new object[2];
            paramsArray[0] = newSh;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("SheetTableUpdate", ref paramsArray);
        }

        public void ModelChange([In, MarshalAs(UnmanagedType.IDispatch)] object changes)
        {
            if (!Validate("ModelChange"))
            {
                Invoker.ReleaseParamsArray(changes);
                return;
            }

            NetOffice.ExcelApi.ModelChanges newChanges = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ModelChanges>(EventClass, changes, typeof(NetOffice.ExcelApi.ModelChanges));
            object[] paramsArray = new object[1];
            paramsArray[0] = newChanges;
            EventBinding.RaiseCustomEvent("ModelChange", ref paramsArray);
        }

        public void SheetBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("SheetBeforeDelete"))
            {
                Invoker.ReleaseParamsArray(sh);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSh;
            EventBinding.RaiseCustomEvent("SheetBeforeDelete", ref paramsArray);
        }

        #endregion
    }
}
