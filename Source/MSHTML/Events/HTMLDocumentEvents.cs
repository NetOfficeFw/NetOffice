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
    [ComImport, Guid("3050F260-98B5-11CF-BB82-00AA00BDCE0B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface HTMLDocumentEvents
	{
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418102)]
		void onhelp();

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

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418103)]
		void onmouseout();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418104)]
		void onmouseover();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-609)]
		void onreadystatechange();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418108)]
		void onbeforeupdate();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418107)]
		void onafterupdate();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418106)]
		void onrowexit();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418105)]
		void onrowenter();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418101)]
		void ondragstart();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418100)]
		void onselectstart();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418099)]
		void onerrorupdate();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1023)]
		void oncontextmenu();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1026)]
		void onstop();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418080)]
		void onrowsdelete();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418079)]
		void onrowsinserted();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418078)]
		void oncellchange();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418093)]
		void onpropertychange();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418098)]
		void ondatasetchanged();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418097)]
		void ondataavailable();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418096)]
		void ondatasetcomplete();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1027)]
		void onbeforeeditfocus();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1037)]
		void onselectionchange();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1036)]
		void oncontrolselect();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1033)]
		void onmousewheel();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1048)]
		void onfocusin();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1049)]
		void onfocusout();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1044)]
		void onactivate();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1045)]
		void ondeactivate();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1047)]
		void onbeforeactivate();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1034)]
		void onbeforedeactivate();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class HTMLDocumentEvents_SinkHelper : SinkHelper, HTMLDocumentEvents
	{
		#region Static
		
		public static readonly string Id = "3050F260-98B5-11CF-BB82-00AA00BDCE0B";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject _eventClass;
        
		#endregion
		
		#region Construction

		public HTMLDocumentEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region HTMLDocumentEvents Members
		
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

		public void onmouseout()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onmouseout");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onmouseout", ref paramsArray);
		}

		public void onmouseover()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onmouseover");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onmouseover", ref paramsArray);
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

		public void onbeforeupdate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onbeforeupdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onbeforeupdate", ref paramsArray);
		}

		public void onafterupdate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onafterupdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onafterupdate", ref paramsArray);
		}

		public void onrowexit()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onrowexit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onrowexit", ref paramsArray);
		}

		public void onrowenter()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onrowenter");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onrowenter", ref paramsArray);
		}

		public void ondragstart()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ondragstart");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("ondragstart", ref paramsArray);
		}

		public void onselectstart()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onselectstart");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onselectstart", ref paramsArray);
		}

		public void onerrorupdate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onerrorupdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onerrorupdate", ref paramsArray);
		}

		public void oncontextmenu()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("oncontextmenu");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("oncontextmenu", ref paramsArray);
		}

		public void onstop()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onstop");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onstop", ref paramsArray);
		}

		public void onrowsdelete()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onrowsdelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onrowsdelete", ref paramsArray);
		}

		public void onrowsinserted()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onrowsinserted");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onrowsinserted", ref paramsArray);
		}

		public void oncellchange()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("oncellchange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("oncellchange", ref paramsArray);
		}

		public void onpropertychange()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onpropertychange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onpropertychange", ref paramsArray);
		}

		public void ondatasetchanged()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ondatasetchanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("ondatasetchanged", ref paramsArray);
		}

		public void ondataavailable()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ondataavailable");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("ondataavailable", ref paramsArray);
		}

		public void ondatasetcomplete()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ondatasetcomplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("ondatasetcomplete", ref paramsArray);
		}

		public void onbeforeeditfocus()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onbeforeeditfocus");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onbeforeeditfocus", ref paramsArray);
		}

		public void onselectionchange()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onselectionchange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onselectionchange", ref paramsArray);
		}

		public void oncontrolselect()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("oncontrolselect");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("oncontrolselect", ref paramsArray);
		}

		public void onmousewheel()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onmousewheel");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onmousewheel", ref paramsArray);
		}

		public void onfocusin()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onfocusin");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onfocusin", ref paramsArray);
		}

		public void onfocusout()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onfocusout");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onfocusout", ref paramsArray);
		}

		public void onactivate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onactivate", ref paramsArray);
		}

		public void ondeactivate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ondeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("ondeactivate", ref paramsArray);
		}

		public void onbeforeactivate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onbeforeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onbeforeactivate", ref paramsArray);
		}

		public void onbeforedeactivate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("onbeforedeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("onbeforedeactivate", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}