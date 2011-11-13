using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.OutlookApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibraryAttribute("Outlook", 9,10,11,12,14)]
	[ComImport, Guid("00063078-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ExplorersEvents
	{
		[SupportByLibraryAttribute("Outlook", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void NewExplorer([In, MarshalAs(UnmanagedType.IDispatch)] object explorer);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ExplorersEvents_SinkHelper : SinkHelper, ExplorersEvents
	{
		#region Static
		
		public static readonly string Id = "00063078-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ExplorersEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region ExplorersEvents Members
		
		public void NewExplorer([In, MarshalAs(UnmanagedType.IDispatch)] object explorer)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewExplorer");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(explorer);
				return;
			}

			NetOffice.OutlookApi._Explorer newExplorer = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, explorer) as NetOffice.OutlookApi._Explorer;
			object[] paramsArray = new object[1];
			paramsArray[0] = newExplorer;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}