using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.MSFormsApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.MSFormsApi.EventContracts.ControlEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ControlEvents_SinkHelper : SinkHelper, NetOffice.MSFormsApi.EventContracts.ControlEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from ControlEvents
		/// </summary>
		public static readonly string Id = "9A4BBF53-4E46-101B-8BBD-00AA003E3B29";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public ControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ControlEvents
		
		/// <summary>
		/// 
		/// </summary>
		public void Enter()
        {
            if (!Validate("Enter"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Enter", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void Exit([In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("Exit"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancel, typeof(NetOffice.MSFormsApi.ReturnBoolean));
			object[] paramsArray = new object[1];
			paramsArray[0] = newCancel;
			EventBinding.RaiseCustomEvent("Exit", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void BeforeUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("BeforeUpdate"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancel, typeof(NetOffice.MSFormsApi.ReturnBoolean));
            object[] paramsArray = new object[1];
			paramsArray[0] = newCancel;
			EventBinding.RaiseCustomEvent("BeforeUpdate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void AfterUpdate()
        {
            if (!Validate("AfterUpdate"))
            {   
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AfterUpdate", ref paramsArray);
		}

		#endregion
	}
	
}

