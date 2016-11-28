using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.MSHTMLApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("MSHTML", 4)]
	[ComImport, Guid("96A0A4E0-D062-11CF-94B6-00AA0060275C"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface HTMLWindowEvents
	{
		[SupportByVersionAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1003)]
		void onload();

		[SupportByVersionAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1008)]
		void onunload();

		[SupportByVersionAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418102)]
		void onhelp();

		[SupportByVersionAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418111)]
		void onfocus();

		[SupportByVersionAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418112)]
		void onblur();

		[SupportByVersionAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1002)]
		void onerror([In] object description, [In] object url, [In] object line);

		[SupportByVersionAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1016)]
		void onresize();

		[SupportByVersionAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1014)]
		void onscroll();

		[SupportByVersionAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1017)]
		void onbeforeunload();

		[SupportByVersionAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1024)]
		void onbeforeprint();

		[SupportByVersionAttribute("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1025)]
		void onafterprint();
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class HTMLWindowEvents_SinkHelper : SinkHelper, HTMLWindowEvents
	{
		#region Static
		
		public static readonly string Id = "96A0A4E0-D062-11CF-94B6-00AA0060275C";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public HTMLWindowEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region Properties

        internal Core Factory
        {
            get
            {
                if (null != _eventClass)
                    return _eventClass.Factory;
                else
                    return Core.Default;
            }
        }

        #endregion

		#region HTMLWindowEvents Members
		
		public void onload()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onload");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onload", ref paramsArray);
		}

		public void onunload()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onunload");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onunload", ref paramsArray);
		}

		public void onhelp()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onhelp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onhelp", ref paramsArray);
		}

		public void onfocus()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onfocus");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onfocus", ref paramsArray);
		}

		public void onblur()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onblur");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onblur", ref paramsArray);
		}

		public void onerror([In] object description, [In] object url, [In] object line)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onerror");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(description, url, line);
				return;
			}

			string newdescription = Convert.ToString(description);
			string newurl = Convert.ToString(url);
			Int32 newline = Convert.ToInt32(line);
			object[] paramsArray = new object[3];
			paramsArray[0] = newdescription;
			paramsArray[1] = newurl;
			paramsArray[2] = newline;
			_eventBinding.RaiseCustomEvent("onerror", ref paramsArray);
		}

		public void onresize()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onresize");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onresize", ref paramsArray);
		}

		public void onscroll()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onscroll");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onscroll", ref paramsArray);
		}

		public void onbeforeunload()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onbeforeunload");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onbeforeunload", ref paramsArray);
		}

		public void onbeforeprint()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onbeforeprint");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onbeforeprint", ref paramsArray);
		}

		public void onafterprint()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onafterprint");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onafterprint", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}