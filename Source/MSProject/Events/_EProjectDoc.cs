using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("MSProject", 11,12,14)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("F81DD3C0-5089-11CF-A49D-00AA00574C74"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _EProjectDoc
	{
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Open([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void BeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void BeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void BeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void Calculate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void Change([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void Activate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void Deactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _EProjectDoc_SinkHelper : SinkHelper, _EProjectDoc
	{
		#region Static
		
		public static readonly string Id = "F81DD3C0-5089-11CF-A49D-00AA00574C74";
		
		#endregion
	
		#region Ctor

		public _EProjectDoc_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _EProjectDoc Members
		
		public void Open([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
        {
            if (!Validate("Open"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("Open", ref paramsArray);
		}

        public void BeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
        {
            if (!Validate("BeforeClose"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("BeforeClose", ref paramsArray);
		}

        public void BeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("BeforeSave"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("BeforeSave", ref paramsArray);
		}

        public void BeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("BeforePrint"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("BeforePrint", ref paramsArray);
		}

        public void Calculate([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("Calculate"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("Calculate", ref paramsArray);
		}

        public void Change([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("Change"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

        public void Activate([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
        {
            if (!Validate("Activate"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("Activate", ref paramsArray);
		}

        public void Deactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("Deactivate"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("Deactivate", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}