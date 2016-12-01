using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.VisioApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[ComImport, Guid("000D0B05-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EStyles
	{
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32772)]
		void StyleAdded([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8196)]
		void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16388)]
		void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(300)]
		void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(301)]
		void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EStyles_SinkHelper : SinkHelper, EStyles
	{
		#region Static
		
		public static readonly string Id = "000D0B05-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public EStyles_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region EStyles Members
		
		public void StyleAdded([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("StyleAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(style);
				return;
			}

			NetOffice.VisioApi.IVStyle newStyle = Factory.CreateObjectFromComProxy(_eventClass, style) as NetOffice.VisioApi.IVStyle;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			_eventBinding.RaiseCustomEvent("StyleAdded", ref paramsArray);
		}

		public void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("StyleChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(style);
				return;
			}

			NetOffice.VisioApi.IVStyle newStyle = Factory.CreateObjectFromComProxy(_eventClass, style) as NetOffice.VisioApi.IVStyle;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			_eventBinding.RaiseCustomEvent("StyleChanged", ref paramsArray);
		}

		public void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeStyleDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(style);
				return;
			}

			NetOffice.VisioApi.IVStyle newStyle = Factory.CreateObjectFromComProxy(_eventClass, style) as NetOffice.VisioApi.IVStyle;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			_eventBinding.RaiseCustomEvent("BeforeStyleDelete", ref paramsArray);
		}

		public void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelStyleDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(style);
				return;
			}

			NetOffice.VisioApi.IVStyle newStyle = Factory.CreateObjectFromComProxy(_eventClass, style) as NetOffice.VisioApi.IVStyle;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			_eventBinding.RaiseCustomEvent("QueryCancelStyleDelete", ref paramsArray);
		}

		public void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("StyleDeleteCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(style);
				return;
			}

			NetOffice.VisioApi.IVStyle newStyle = Factory.CreateObjectFromComProxy(_eventClass, style) as NetOffice.VisioApi.IVStyle;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			_eventBinding.RaiseCustomEvent("StyleDeleteCanceled", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}