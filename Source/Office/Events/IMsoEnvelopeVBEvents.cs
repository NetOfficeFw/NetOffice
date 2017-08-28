using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Office", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000672AD-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface IMsoEnvelopeVBEvents
	{
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void EnvelopeShow();

		[SupportByVersion("Office", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void EnvelopeHide();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class IMsoEnvelopeVBEvents_SinkHelper : SinkHelper, IMsoEnvelopeVBEvents
	{
		#region Static
		
		public static readonly string Id = "000672AD-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject _eventClass;
        
		#endregion
		
		#region Construction

		public IMsoEnvelopeVBEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region IMsoEnvelopeVBEvents Members
		
		public void EnvelopeShow()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("EnvelopeShow");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("EnvelopeShow", ref paramsArray);
		}

		public void EnvelopeHide()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("EnvelopeHide");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("EnvelopeHide", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}