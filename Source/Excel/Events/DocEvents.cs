using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.Events
{
    #pragma warning disable

    #region SinkPoint Interface

    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00024411-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface DocEvents
    {
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1543)]
        void SelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1537)]
        void BeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1534)]
        void BeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
        void Activate();

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
        void Deactivate();

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(279)]
        void Calculate();

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("target", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1545)]
        void Change([In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("target", typeof(ExcelApi.Hyperlink))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1470)]
        void FollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2156)]
        void PivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("targetRange", typeof(ExcelApi.Range))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2886)]
        void PivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2889)]
        void PivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2892)]
        void PivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("targetPivotTable", typeof(ExcelApi.PivotTable))]
        [SinkArgument("valueChangeStart", SinkArgumentType.Int32)]
        [SinkArgument("valueChangeEnd", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2893)]
        void PivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd);

        [SupportByVersion("Excel", 14, 15, 16)]
        [SinkArgument("target", typeof(ExcelApi.PivotTable))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2894)]
        void PivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3070)]
        void LensGalleryRenderComplete();

        [SupportByVersion("Excel", 15, 16)]
        [SinkArgument("target", typeof(ExcelApi.TableObject))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3071)]
        void TableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object target);

        [SupportByVersion("Excel", 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3072)]
        void BeforeDelete();
    }

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class DocEvents_SinkHelper : SinkHelper, DocEvents
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

            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
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

            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
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

            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
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

            NetOffice.ExcelApi.Range newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, target, NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
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

            NetOffice.ExcelApi.Hyperlink newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Hyperlink>(EventClass, target, NetOffice.ExcelApi.Hyperlink.LateBindingApiWrapperType);
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

            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
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

            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
            NetOffice.ExcelApi.Range newTargetRange = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.Range>(EventClass, targetRange, NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
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

            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
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

            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
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

            NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, targetPivotTable, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
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

            NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.PivotTable>(EventClass, target, NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
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

            NetOffice.ExcelApi.TableObject newTarget = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.TableObject>(EventClass, target, NetOffice.ExcelApi.TableObject.LateBindingApiWrapperType);
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

    #endregion

    #pragma warning restore
}