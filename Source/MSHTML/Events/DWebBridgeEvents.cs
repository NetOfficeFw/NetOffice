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
    [ComImport, Guid("A6D897FF-0A95-11D1-B0BA-006008166E11"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DWebBridgeEvents
	{
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void onscriptletevent([In] object name, [In] object eventData);

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-609)]
		void onreadystatechange();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void onclick();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void ondblclick();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void onkeydown();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void onkeyup();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void onkeypress();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void onmousedown();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void onmousemove();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void onmouseup();
	}
	
	#endregion
	
	#region SinkHelper
	
    [InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DWebBridgeEvents_SinkHelper : SinkHelper, DWebBridgeEvents
	{
		#region Static
		
		public static readonly string Id = "A6D897FF-0A95-11D1-B0BA-006008166E11";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public DWebBridgeEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region DWebBridgeEvents Members
		
		public void onscriptletevent([In] object name, [In] object eventData)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onscriptletevent");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(name, eventData);
				return;
			}

			string newname = Convert.ToString(name);
			object neweventData = (object)eventData;
			object[] paramsArray = new object[2];
			paramsArray[0] = newname;
			paramsArray[1] = neweventData;
			_eventBinding.RaiseCustomEvent("onscriptletevent", ref paramsArray);
		}

		public void onreadystatechange()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onreadystatechange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onreadystatechange", ref paramsArray);
		}

		public void onclick()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onclick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onclick", ref paramsArray);
		}

		public void ondblclick()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ondblclick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("ondblclick", ref paramsArray);
		}

		public void onkeydown()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onkeydown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onkeydown", ref paramsArray);
		}

		public void onkeyup()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onkeyup");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onkeyup", ref paramsArray);
		}

		public void onkeypress()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onkeypress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onkeypress", ref paramsArray);
		}

		public void onmousedown()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onmousedown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onmousedown", ref paramsArray);
		}

		public void onmousemove()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onmousemove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onmousemove", ref paramsArray);
		}

		public void onmouseup()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onmouseup");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onmouseup", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}