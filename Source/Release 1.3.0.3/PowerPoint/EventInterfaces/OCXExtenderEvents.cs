using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.PowerPointApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibraryAttribute("PowerPoint", 9,10,11,12,14)]
	[ComImport, Guid("914934C1-5A91-11CF-8700-00AA0060263B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OCXExtenderEvents
	{
		[SupportByLibraryAttribute("PowerPoint", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147417888)]
		void GotFocus();

		[SupportByLibraryAttribute("PowerPoint", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147417887)]
		void LostFocus();
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OCXExtenderEvents_SinkHelper : SinkHelper, OCXExtenderEvents
	{
		#region Static
		
		public static readonly string Id = "914934C1-5A91-11CF-8700-00AA0060263B";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public OCXExtenderEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region OCXExtenderEvents Members
		
		public void GotFocus()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("GotFocus");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void LostFocus()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("LostFocus");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}