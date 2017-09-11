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
		
		#region Ctor

		public IMsoEnvelopeVBEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region IMsoEnvelopeVBEvents
		
		public void EnvelopeShow()
        {
            if (!Validate("EnvelopeShow"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("EnvelopeShow", ref paramsArray);
		}

		public void EnvelopeHide()
        {
            if (!Validate("EnvelopeHide"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("EnvelopeHide", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}