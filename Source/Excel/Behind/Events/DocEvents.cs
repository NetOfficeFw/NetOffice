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
    public class DocEvents_SinkHelper : SinkHelper, NetOffice.ExcelApi.EventContracts.DocEvents
    {
        #region Static

        public static readonly string Id = "00024411-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public DocEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region DocEvents

        public void SelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("SelectionChange"))
            {
                Invoker.ReleaseParamsArray(target);
                return;
            }

            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, typeof(NetOffice.ExcelApi.Range));
            object[] paramsArray = new object[1];
            paramsArray[0] = newTarget;
            EventBinding.RaiseCustomEvent("SelectionChange", ref paramsArray);
        }

        public void BeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeDoubleClick"))
            {
                Invoker.ReleaseParamsArray(target, cancel);
                return;
            }

            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, typeof(NetOffice.ExcelApi.Range));
            object[] paramsArray = new object[2];
            paramsArray[0] = newTarget;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("BeforeDoubleClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void BeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeRightClick"))
            {
                Invoker.ReleaseParamsArray(target, cancel);
                return;
            }

            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, typeof(NetOffice.ExcelApi.Range));
            object[] paramsArray = new object[2];
            paramsArray[0] = newTarget;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("BeforeRightClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
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

        public void Calculate()
        {
            if (!Validate("Calculate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Calculate", ref paramsArray);
        }

        public void Change([In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("Change"))
            {
                Invoker.ReleaseParamsArray(target);
                return;
            }

            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, typeof(NetOffice.ExcelApi.Range));
            object[] paramsArray = new object[1];
            paramsArray[0] = newTarget;
            EventBinding.RaiseCustomEvent("Change", ref paramsArray);
        }

        public void FollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("FollowHyperlink"))
            {
                Invoker.ReleaseParamsArray(target);
                return;
            }

            NetOffice.ExcelApi.Hyperlink newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Hyperlink>(EventClass, target, typeof(NetOffice.ExcelApi.Hyperlink));
            object[] paramsArray = new object[1];
            paramsArray[0] = newTarget;
            EventBinding.RaiseCustomEvent("FollowHyperlink", ref paramsArray);
        }

        public void PivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("PivotTableUpdate"))
            {
                Invoker.ReleaseParamsArray(target);
                return;
            }

            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, typeof(NetOffice.ExcelApi.PivotTable));
            object[] paramsArray = new object[1];
            paramsArray[0] = newTarget;
            EventBinding.RaiseCustomEvent("PivotTableUpdate", ref paramsArray);
        }

        public void PivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange)
        {
            if (!Validate("PivotTableAfterValueChange"))
            {
                Invoker.ReleaseParamsArray(targetPivotTable, targetRange);
                return;
            }

            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, typeof(NetOffice.ExcelApi.PivotTable));
            NetOffice.ExcelApi.Range newTargetRange = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, targetRange, typeof(NetOffice.ExcelApi.Range));
            object[] paramsArray = new object[2];
            paramsArray[0] = newTargetPivotTable;
            paramsArray[1] = newTargetRange;
            EventBinding.RaiseCustomEvent("PivotTableAfterValueChange", ref paramsArray);
        }

        public void PivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
        {
            if (!Validate("PivotTableBeforeAllocateChanges"))
            {
                Invoker.ReleaseParamsArray(targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
                return;
            }

            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, typeof(NetOffice.ExcelApi.PivotTable));
            Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
            Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
            object[] paramsArray = new object[4];
            paramsArray[0] = newTargetPivotTable;
            paramsArray[1] = newValueChangeStart;
            paramsArray[2] = newValueChangeEnd;
            paramsArray.SetValue(cancel, 3);
            EventBinding.RaiseCustomEvent("PivotTableBeforeAllocateChanges", ref paramsArray);

            cancel = ToBoolean(paramsArray[3]);
        }

        public void PivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
        {
            if (!Validate("PivotTableBeforeCommitChanges"))
            {
                Invoker.ReleaseParamsArray(targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
                return;
            }

            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, typeof(NetOffice.ExcelApi.PivotTable));
            Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
            Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
            object[] paramsArray = new object[4];
            paramsArray[0] = newTargetPivotTable;
            paramsArray[1] = newValueChangeStart;
            paramsArray[2] = newValueChangeEnd;
            paramsArray.SetValue(cancel, 3);
            EventBinding.RaiseCustomEvent("PivotTableBeforeCommitChanges", ref paramsArray);

            cancel = ToBoolean(paramsArray[3]);
        }

        public void PivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd)
        {
            if (!Validate("PivotTableBeforeDiscardChanges"))
            {
                Invoker.ReleaseParamsArray(targetPivotTable, valueChangeStart, valueChangeEnd);
                return;
            }

            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, typeof(NetOffice.ExcelApi.PivotTable));
            Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
            Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
            object[] paramsArray = new object[3];
            paramsArray[0] = newTargetPivotTable;
            paramsArray[1] = newValueChangeStart;
            paramsArray[2] = newValueChangeEnd;
            EventBinding.RaiseCustomEvent("PivotTableBeforeDiscardChanges", ref paramsArray);
        }

        public void PivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("PivotTableChangeSync"))
            {
                Invoker.ReleaseParamsArray(target);
                return;
            }

            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, typeof(NetOffice.ExcelApi.PivotTable));
            object[] paramsArray = new object[1];
            paramsArray[0] = newTarget;
            EventBinding.RaiseCustomEvent("PivotTableChangeSync", ref paramsArray);
        }

        public void LensGalleryRenderComplete()
        {
            if (!Validate("LensGalleryRenderComplete"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("LensGalleryRenderComplete", ref paramsArray);
        }

        public void TableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object target)
        {
            if (!Validate("TableUpdate"))
            {
                Invoker.ReleaseParamsArray(target);
                return;
            }

            NetOffice.ExcelApi.TableObject newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.TableObject>(EventClass, target, typeof(NetOffice.ExcelApi.TableObject));
            object[] paramsArray = new object[1];
            paramsArray[0] = newTarget;
            EventBinding.RaiseCustomEvent("TableUpdate", ref paramsArray);
        }

        public void BeforeDelete()
        {
            if (!Validate("BeforeDelete"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("BeforeDelete", ref paramsArray);
        }

        #endregion
    }
}
