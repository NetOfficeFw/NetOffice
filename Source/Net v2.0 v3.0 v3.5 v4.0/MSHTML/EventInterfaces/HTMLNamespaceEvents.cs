using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.MSHTMLApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibraryAttribute("MSHTML", 4)]
	[ComImport, Guid("3050F6BD-98B5-11CF-BB82-00AA00BDCE0B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface HTMLNamespaceEvents
	{
		[SupportByLibraryAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-609)]
		void onreadystatechange([In, MarshalAs(UnmanagedType.IDispatch)] object pEvtObj);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class HTMLNamespaceEvents_SinkHelper : SinkHelper, HTMLNamespaceEvents
	{
		#region Static
		
		public static readonly string Id = "3050F6BD-98B5-11CF-BB82-00AA00BDCE0B";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public HTMLNamespaceEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region HTMLNamespaceEvents Members
		
		public void onreadystatechange([In, MarshalAs(UnmanagedType.IDispatch)] object pEvtObj)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onreadystatechange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pEvtObj);
				return;
			}

			LateBindingApi.MSHTMLApi.IHTMLEventObj newpEvtObj = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pEvtObj) as LateBindingApi.MSHTMLApi.IHTMLEventObj;
			object[] paramsArray = new object[1];
			paramsArray[0] = newpEvtObj;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}