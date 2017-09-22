using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("OWC10", 1)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("F5B39A9B-1480-11D3-8549-00C04FAC67D7"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _DataSourceControlEvent
	{
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(624)]
		void Current([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(626)]
		void BeforeExpand([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(627)]
		void BeforeCollapse([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(628)]
		void BeforeFirstPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(629)]
		void BeforePreviousPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(630)]
		void BeforeNextPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(631)]
		void BeforeLastPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(632)]
		void DataError([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(633)]
		void DataPageComplete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(634)]
		void BeforeInitialBind([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(635)]
		void RecordsetSaveProgress([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(636)]
		void AfterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(637)]
		void AfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(638)]
		void AfterUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(639)]
		void BeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(640)]
		void BeforeInsert([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(641)]
		void BeforeOverwrite([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(642)]
		void BeforeUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(643)]
		void Dirty([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(644)]
		void RecordExit([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(647)]
		void Undo([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(648)]
		void Focus([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);
	}
	
	#endregion
	
	#region SinkHelper
	
    [InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _DataSourceControlEvent_SinkHelper : SinkHelper, _DataSourceControlEvent
	{
		#region Static
		
		public static readonly string Id = "F5B39A9B-1480-11D3-8549-00C04FAC67D7";
		
		#endregion
	
		#region Ctor

		public _DataSourceControlEvent_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion		

		#region _DataSourceControlEvent
		
		public void Current([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("Current"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

			NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("Current", ref paramsArray);
		}

		public void BeforeExpand([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("BeforeExpand"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("BeforeExpand", ref paramsArray);
		}

		public void BeforeCollapse([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("BeforeCollapse"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("BeforeCollapse", ref paramsArray);
		}

		public void BeforeFirstPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("BeforeFirstPage"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("BeforeFirstPage", ref paramsArray);
		}

		public void BeforePreviousPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforePreviousPage"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("BeforePreviousPage", ref paramsArray);
		}

		public void BeforeNextPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("BeforeNextPage"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("BeforeNextPage", ref paramsArray);
		}

		public void BeforeLastPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("BeforeLastPage"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("BeforeLastPage", ref paramsArray);
		}

		public void DataError([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("DataError"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("DataError", ref paramsArray);
		}

		public void DataPageComplete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("DataPageComplete"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("DataPageComplete", ref paramsArray);
		}

		public void BeforeInitialBind([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("BeforeInitialBind"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("BeforeInitialBind", ref paramsArray);
		}

		public void RecordsetSaveProgress([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("RecordsetSaveProgress"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("RecordsetSaveProgress", ref paramsArray);
		}

		public void AfterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("AfterDelete"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("AfterDelete", ref paramsArray);
		}

		public void AfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("AfterInsert"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("AfterInsert", ref paramsArray);
		}

		public void AfterUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("AfterUpdate"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("AfterUpdate", ref paramsArray);
		}

		public void BeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("BeforeDelete"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("BeforeDelete", ref paramsArray);
		}

		public void BeforeInsert([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("BeforeInsert"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("BeforeInsert", ref paramsArray);
		}

		public void BeforeOverwrite([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("BeforeOverwrite"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("BeforeOverwrite", ref paramsArray);
		}

		public void BeforeUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("BeforeUpdate"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("BeforeUpdate", ref paramsArray);
		}

		public void Dirty([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("Dirty"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("Dirty", ref paramsArray);
		}

		public void RecordExit([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("RecordExit"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("RecordExit", ref paramsArray);
		}

		public void Undo([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("Undo"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("Undo", ref paramsArray);
		}

		public void Focus([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
		{
            if (!Validate("Focus"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, NetOffice.OWC10Api.DSCEventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDSCEventInfo;
			EventBinding.RaiseCustomEvent("Focus", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}