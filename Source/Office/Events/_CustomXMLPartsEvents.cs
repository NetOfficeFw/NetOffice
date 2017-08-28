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

	[SupportByVersion("Office", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000CDB0B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomXMLPartsEvents
	{
		[SupportByVersion("Office", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void PartAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newPart);

		[SupportByVersion("Office", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void PartBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldPart);

		[SupportByVersion("Office", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void PartAfterLoad([In, MarshalAs(UnmanagedType.IDispatch)] object part);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CustomXMLPartsEvents_SinkHelper : SinkHelper, _CustomXMLPartsEvents
	{
		#region Static
		
		public static readonly string Id = "000CDB0B-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject _eventClass;
        
		#endregion
		
		#region Construction

		public _CustomXMLPartsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region _CustomXMLPartsEvents Members
		
		public void PartAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newPart)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PartAfterAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(newPart);
				return;
			}

			NetOffice.OfficeApi.CustomXMLPart newNewPart = Factory.CreateObjectFromComProxy(_eventClass, newPart) as NetOffice.OfficeApi.CustomXMLPart;
			object[] paramsArray = new object[1];
			paramsArray[0] = newNewPart;
			_eventBinding.RaiseCustomEvent("PartAfterAdd", ref paramsArray);
		}

		public void PartBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldPart)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PartBeforeDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(oldPart);
				return;
			}

			NetOffice.OfficeApi.CustomXMLPart newOldPart = Factory.CreateObjectFromComProxy(_eventClass, oldPart) as NetOffice.OfficeApi.CustomXMLPart;
			object[] paramsArray = new object[1];
			paramsArray[0] = newOldPart;
			_eventBinding.RaiseCustomEvent("PartBeforeDelete", ref paramsArray);
		}

		public void PartAfterLoad([In, MarshalAs(UnmanagedType.IDispatch)] object part)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PartAfterLoad");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(part);
				return;
			}

			NetOffice.OfficeApi.CustomXMLPart newPart = Factory.CreateObjectFromComProxy(_eventClass, part) as NetOffice.OfficeApi.CustomXMLPart;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPart;
			_eventBinding.RaiseCustomEvent("PartAfterLoad", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}