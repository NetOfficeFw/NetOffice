using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.MSProjectApi.Behind.EventContracts
{

    /// <summary>
    /// Default implementation of <see cref="NetOffice.MSProjectApi.EventContracts._EProjectDoc"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _EProjectDoc_SinkHelper : SinkHelper, NetOffice.MSProjectApi.EventContracts._EProjectDoc
	{
		#region Static
		
		/// <summary>
		/// Interface Id from _EProjectDoc
		/// </summary>
		public static readonly string Id = "F81DD3C0-5089-11CF-A49D-00AA00574C74";
		
		#endregion
	
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public _EProjectDoc_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _EProjectDoc Members
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="pj"></param>
		public void Open([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
        {
            if (!Validate("Open"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, typeof(NetOffice.MSProjectApi.Project));
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("Open", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="pj"></param>
        public void BeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
        {
            if (!Validate("BeforeClose"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, typeof(NetOffice.MSProjectApi.Project));
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("BeforeClose", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="pj"></param>
        public void BeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("BeforeSave"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, typeof(NetOffice.MSProjectApi.Project));
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("BeforeSave", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="pj"></param>
        public void BeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("BeforePrint"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, typeof(NetOffice.MSProjectApi.Project));
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("BeforePrint", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="pj"></param>
        public void Calculate([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("Calculate"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, typeof(NetOffice.MSProjectApi.Project));
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("Calculate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="pj"></param>
        public void Change([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("Change"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, typeof(NetOffice.MSProjectApi.Project));
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="pj"></param>
        public void Activate([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
        {
            if (!Validate("Activate"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, typeof(NetOffice.MSProjectApi.Project));
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("Activate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="pj"></param>
        public void Deactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("Deactivate"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, typeof(NetOffice.MSProjectApi.Project));
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("Deactivate", ref paramsArray);
		}

		#endregion
	}
	
}

