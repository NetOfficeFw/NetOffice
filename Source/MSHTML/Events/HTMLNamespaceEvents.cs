using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("MSHTML", 4)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("3050F6BD-98B5-11CF-BB82-00AA00BDCE0B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface HTMLNamespaceEvents
	{
		[SupportByVersion("MSHTML", 4)]
        [SinkArgument("pEvtObj", typeof(MSHTMLApi.IHTMLEventObj))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-609)]
		void onreadystatechange([In, MarshalAs(UnmanagedType.IDispatch)] object pEvtObj);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class HTMLNamespaceEvents_SinkHelper : SinkHelper, HTMLNamespaceEvents
	{
		#region Static
		
		public static readonly string Id = "3050F6BD-98B5-11CF-BB82-00AA00BDCE0B";
		
		#endregion
	
		#region Fields

		private IEventBinding	EventBinding;
        private ICOMObject EventClass;
        
		#endregion
		
		#region Ctor

		public HTMLNamespaceEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region HTMLNamespaceEvents
 
		public void onreadystatechange([In, MarshalAs(UnmanagedType.IDispatch)] object pEvtObj)
		{
            if (!Validate("onreadystatechange"))
            {
                return;
            }

			NetOffice.MSHTMLApi.IHTMLEventObj newpEvtObj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLEventObj>(EventClass, pEvtObj, NetOffice.MSHTMLApi.IHTMLEventObj.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newpEvtObj;
			EventBinding.RaiseCustomEvent("onreadystatechange", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}