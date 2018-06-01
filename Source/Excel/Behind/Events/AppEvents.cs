using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Exceptions;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.ExcelApi.EventContracts.AppEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class AppEvents_SinkHelper : SinkHelper, NetOffice.ExcelApi.EventContracts.AppEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from AppEvents
        /// </summary>
        public static readonly string Id = "00024413-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public AppEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region AppEvents

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        public void NewWorkbook([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("NewWorkbook"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("NewWorkbook", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        /// <param name="cancel"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        /// <param name="cancel"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        public void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SheetChange"))
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        public void WorkbookOpen([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("WorkbookOpen"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("WorkbookOpen", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        public void WorkbookActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("WorkbookActivate"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("WorkbookActivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        public void WorkbookDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("WorkbookDeactivate"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("WorkbookDeactivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="cancel"></param>
        public void WorkbookBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel)
        {
            if (!Validate("WorkbookBeforeClose"))
            {
                Invoker.ReleaseParamsArray(wb, cancel);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("WorkbookBeforeClose", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="saveAsUI"></param>
        /// <param name="cancel"></param>
        public void WorkbookBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object saveAsUI, [In] [Out] ref object cancel)
        {
            if (!Validate("WorkbookBeforeSave"))
            {
                Invoker.ReleaseParamsArray(wb, saveAsUI, cancel);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            bool newSaveAsUI = ToBoolean(saveAsUI);
            object[] paramsArray = new object[3];
            paramsArray[0] = newWb;
            paramsArray[1] = newSaveAsUI;
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("WorkbookBeforeSave", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="cancel"></param>
        public void WorkbookBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel)
        {
            if (!Validate("WorkbookBeforeSave"))
            {
                Invoker.ReleaseParamsArray(wb, cancel);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("WorkbookBeforeSave", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sh"></param>
        public void WorkbookNewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (!Validate("WorkbookNewSheet"))
            {
                Invoker.ReleaseParamsArray(wb, sh);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newSh;
            EventBinding.RaiseCustomEvent("WorkbookNewSheet", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        public void WorkbookAddinInstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("WorkbookAddinInstall"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("WorkbookAddinInstall", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        public void WorkbookAddinUninstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
        {
            if (!Validate("WorkbookAddinUninstall"))
            {
                Invoker.ReleaseParamsArray(wb);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWb;
            EventBinding.RaiseCustomEvent("WorkbookAddinUninstall", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="wn"></param>
        public void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowResize"))
            {
                Invoker.ReleaseParamsArray(wb, wn);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.ExcelApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Window>(EventClass, wb, typeof(NetOffice.ExcelApi.Window));
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowResize", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="wn"></param>
        public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowActivate"))
            {
                Invoker.ReleaseParamsArray(wb, wn);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.ExcelApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Window>(EventClass, wb, typeof(NetOffice.ExcelApi.Window));
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="wn"></param>
        public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(wb, wn);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.ExcelApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Window>(EventClass, wb, typeof(NetOffice.ExcelApi.Window));
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="target"></param>
        public void WorkbookPivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("WorkbookPivotTableCloseConnection"))
            {
                Invoker.ReleaseParamsArray(wb, target);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, typeof(NetOffice.ExcelApi.PivotTable));
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("WorkbookPivotTableCloseConnection", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="target"></param>
        public void WorkbookPivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("WorkbookPivotTableOpenConnection"))
            {
                Invoker.ReleaseParamsArray(wb, target);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, typeof(NetOffice.ExcelApi.PivotTable));
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newTarget;
            EventBinding.RaiseCustomEvent("WorkbookPivotTableOpenConnection", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="syncEventType"></param>
        public void WorkbookSync([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object syncEventType)
        {
            if (!Validate("WorkbookSync"))
            {
                Invoker.ReleaseParamsArray(wb, syncEventType);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newSyncEventType;
            EventBinding.RaiseCustomEvent("WorkbookSync", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="map"></param>
        /// <param name="url"></param>
        /// <param name="isRefresh"></param>
        /// <param name="cancel"></param>
        public void WorkbookBeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel)
        {
            if (!Validate("WorkbookBeforeXmlImport"))
            {
                Invoker.ReleaseParamsArray(wb, map, url, isRefresh, cancel);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, typeof(NetOffice.ExcelApi.XmlMap));
            string newUrl = Convert.ToString(url);
            bool newIsRefresh = ToBoolean(isRefresh);
            object[] paramsArray = new object[5];
            paramsArray[0] = newWb;
            paramsArray[1] = newMap;
            paramsArray[2] = newUrl;
            paramsArray[3] = newIsRefresh;
            paramsArray.SetValue(cancel, 4);
            EventBinding.RaiseCustomEvent("WorkbookBeforeXmlImport", ref paramsArray);

            cancel = ToBoolean(paramsArray[4]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="map"></param>
        /// <param name="isRefresh"></param>
        /// <param name="result"></param>
        public void WorkbookAfterXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object isRefresh, [In] object result)
        {
            if (!Validate("WorkbookAfterXmlImport"))
            {
                Invoker.ReleaseParamsArray(wb, map, isRefresh, result);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, typeof(NetOffice.ExcelApi.XmlMap));
            bool newIsRefresh = ToBoolean(isRefresh);
            NetOffice.ExcelApi.Enums.XlXmlImportResult newResult = (NetOffice.ExcelApi.Enums.XlXmlImportResult)result;
            object[] paramsArray = new object[4];
            paramsArray[0] = newWb;
            paramsArray[1] = newMap;
            paramsArray[2] = newIsRefresh;
            paramsArray[3] = newResult;
            EventBinding.RaiseCustomEvent("WorkbookAfterXmlImport", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="map"></param>
        /// <param name="url"></param>
        /// <param name="cancel"></param>
        public void WorkbookBeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel)
        {
            if (!Validate("WorkbookBeforeXmlExport"))
            {
                Invoker.ReleaseParamsArray(wb, map, url, cancel);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, typeof(NetOffice.ExcelApi.XmlMap));
            string newUrl = ToString(url);
            object[] paramsArray = new object[4];
            paramsArray[0] = newWb;
            paramsArray[1] = newMap;
            paramsArray[2] = newUrl;
            paramsArray.SetValue(cancel, 3);
            EventBinding.RaiseCustomEvent("WorkbookBeforeXmlExport", ref paramsArray);

            cancel = ToBoolean(paramsArray[3]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="map"></param>
        /// <param name="url"></param>
        /// <param name="result"></param>
        public void WorkbookAfterXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object result)
        {
            if (!Validate("WorkbookAfterXmlExport"))
            {
                Invoker.ReleaseParamsArray(wb, map, url, result);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.ExcelApi.XmlMap newMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(EventClass, map, typeof(NetOffice.ExcelApi.XmlMap));
            string newUrl = ToString(url);
            NetOffice.ExcelApi.Enums.XlXmlExportResult newResult = (NetOffice.ExcelApi.Enums.XlXmlExportResult)result;
            object[] paramsArray = new object[4];
            paramsArray[0] = newWb;
            paramsArray[1] = newMap;
            paramsArray[2] = newUrl;
            paramsArray[3] = newResult;
            EventBinding.RaiseCustomEvent("WorkbookAfterXmlExport", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="description"></param>
        /// <param name="sheet"></param>
        /// <param name="success"></param>
        public void WorkbookRowsetComplete([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object description, [In] object sheet, [In] object success)
        {
            if (!Validate("WorkbookRowsetComplete"))
            {
                Invoker.ReleaseParamsArray(wb, description, sheet, success);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            string newDescription = ToString(description);
            string newSheet = ToString(sheet);
            bool newSuccess = ToBoolean(success);
            object[] paramsArray = new object[4];
            paramsArray[0] = newWb;
            paramsArray[1] = newDescription;
            paramsArray[2] = newSheet;
            paramsArray[3] = newSuccess;
            EventBinding.RaiseCustomEvent("WorkbookRowsetComplete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        public void AfterCalculate()
        {
            if (!Validate("AfterCalculate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("AfterCalculate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="targetPivotTable"></param>
        /// <param name="targetRange"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="targetPivotTable"></param>
        /// <param name="valueChangeStart"></param>
        /// <param name="valueChangeEnd"></param>
        /// <param name="cancel"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="targetPivotTable"></param>
        /// <param name="valueChangeStart"></param>
        /// <param name="valueChangeEnd"></param>
        /// <param name="cancel"></param>
        public void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
        {
            if (!Validate("SheetPivotTableBeforeCommitChanges"))
            {
                Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, typeof(NetOffice.ExcelApi.PivotTable));
            Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
            Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
            object[] paramsArray = new object[5];
            paramsArray[0] = newSh;
            paramsArray[1] = newTargetPivotTable;
            paramsArray[2] = newValueChangeStart;
            paramsArray[3] = newValueChangeEnd;
            paramsArray.SetValue(cancel, 4);
            EventBinding.RaiseCustomEvent("SheetPivotTableBeforeCommitChanges", ref paramsArray);

            cancel = ToBoolean(paramsArray[4]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="targetPivotTable"></param>
        /// <param name="valueChangeStart"></param>
        /// <param name="valueChangeEnd"></param>
        public void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd)
        {
            if (!Validate("SheetPivotTableBeforeDiscardChanges"))
            {
                Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd);
                return;
            }

            object newSh = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sh) as object;
            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, typeof(NetOffice.ExcelApi.PivotTable));
            Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
            Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
            object[] paramsArray = new object[4];
            paramsArray[0] = newSh;
            paramsArray[1] = newTargetPivotTable;
            paramsArray[2] = newValueChangeStart;
            paramsArray[3] = newValueChangeEnd;
            EventBinding.RaiseCustomEvent("SheetPivotTableBeforeDiscardChanges", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pvw"></param>
        public void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
        {
            if (!Validate("ProtectedViewWindowOpen"))
            {
                Invoker.ReleaseParamsArray(pvw);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, typeof(NetOffice.ExcelApi.ProtectedViewWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPvw;
            EventBinding.RaiseCustomEvent("ProtectedViewWindowOpen", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pvw"></param>
        /// <param name="cancel"></param>
        public void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] [Out] ref object cancel)
        {
            if (!Validate("ProtectedViewWindowBeforeEdit"))
            {
                Invoker.ReleaseParamsArray(pvw, cancel);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, typeof(NetOffice.ExcelApi.ProtectedViewWindow));
            object[] paramsArray = new object[2];
            paramsArray[0] = newPvw;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeEdit", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pvw"></param>
        /// <param name="reason"></param>
        /// <param name="cancel"></param>
        public void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] object reason, [In] [Out] ref object cancel)
        {
            if (!Validate("ProtectedViewWindowBeforeClose"))
            {
                Invoker.ReleaseParamsArray(pvw, reason, cancel);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, typeof(NetOffice.ExcelApi.ProtectedViewWindow));
            NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason newReason = (NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason)reason;
            object[] paramsArray = new object[3];
            paramsArray[0] = newPvw;
            paramsArray[1] = newReason;
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeClose", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pvw"></param>
        public void ProtectedViewWindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
        {
            if (!Validate("ProtectedViewWindowResize"))
            {
                Invoker.ReleaseParamsArray(pvw);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, typeof(NetOffice.ExcelApi.ProtectedViewWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPvw;
            EventBinding.RaiseCustomEvent("ProtectedViewWindowResize", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pvw"></param>
        public void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
        {
            if (!Validate("ProtectedViewWindowActivate"))
            {
                Invoker.ReleaseParamsArray(pvw);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, typeof(NetOffice.ExcelApi.ProtectedViewWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPvw;
            EventBinding.RaiseCustomEvent("ProtectedViewWindowActivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pvw"></param>
        public void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
        {
            if (!Validate("ProtectedViewWindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(pvw);
                return;
            }

            NetOffice.ExcelApi.ProtectedViewWindow newPvw = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ProtectedViewWindow>(EventClass, pvw, typeof(NetOffice.ExcelApi.ProtectedViewWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPvw;
            EventBinding.RaiseCustomEvent("ProtectedViewWindowDeactivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="success"></param>
        public void WorkbookAfterSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object success)
        {
            if (!Validate("WorkbookAfterSave"))
            {
                Invoker.ReleaseParamsArray(wb, success);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            bool newSuccess = ToBoolean(success);
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newSuccess;
            EventBinding.RaiseCustomEvent("WorkbookAfterSave", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="ch"></param>
        public void WorkbookNewChart([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object ch)
        {
            if (!Validate("WorkbookNewChart"))
            {
                Invoker.ReleaseParamsArray(wb, ch);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.ExcelApi.Chart newCh = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Chart>(EventClass, ch, typeof(NetOffice.ExcelApi.Chart));
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newCh;
            EventBinding.RaiseCustomEvent("WorkbookNewChart", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="changes"></param>
        public void WorkbookModelChange([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object changes)
        {
            if (!Validate("WorkbookModelChange"))
            {
                Invoker.ReleaseParamsArray(wb, changes);
                return;
            }

            NetOffice.ExcelApi.Workbook newWb = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Workbook>(EventClass, wb, typeof(NetOffice.ExcelApi.Workbook));
            NetOffice.ExcelApi.ModelChanges newChanges = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.ModelChanges>(EventClass, changes, typeof(NetOffice.ExcelApi.ModelChanges));
            object[] paramsArray = new object[2];
            paramsArray[0] = newWb;
            paramsArray[1] = newChanges;
            EventBinding.RaiseCustomEvent("WorkbookModelChange", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sh"></param>
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
