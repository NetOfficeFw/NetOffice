using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00063085-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface SyncObjectEvents
	{
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void SyncStart();

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("state", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlSyncState))]
        [SinkArgument("description", SinkArgumentType.String)]
        [SinkArgument("value", SinkArgumentType.Int32)]
        [SinkArgument("max", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void Progress([In] object state, [In] object description, [In] object value, [In] object max);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int32)]
        [SinkArgument("description", SinkArgumentType.String)]       
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void OnError([In] object code, [In] object description);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void SyncEnd();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class SyncObjectEvents_SinkHelper : SinkHelper, SyncObjectEvents
	{
		#region Static
		
		public static readonly string Id = "00063085-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public SyncObjectEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region SyncObjectEvents
		
		public void SyncStart()
		{
            if (!Validate("SyncStart"))
            {              
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("SyncStart", ref paramsArray);
		}

        public void Progress([In] object state, [In] object description, [In] object value, [In] object max)
        {
            if (!Validate("Progress"))
            {
                Invoker.ReleaseParamsArray(state, description, value, max);
                return;
            }

			NetOffice.OutlookApi.Enums.OlSyncState newState = (NetOffice.OutlookApi.Enums.OlSyncState)state;
			string newDescription = ToString(description);
			Int32 newValue = ToInt32(value);
			Int32 newMax = ToInt32(max);
			object[] paramsArray = new object[4];
			paramsArray[0] = newState;
			paramsArray[1] = newDescription;
			paramsArray[2] = newValue;
			paramsArray[3] = newMax;
			EventBinding.RaiseCustomEvent("Progress", ref paramsArray);
		}

		public void OnError([In] object code, [In] object description)
        {
            if (!Validate("OnError"))
            {
                Invoker.ReleaseParamsArray(code, description);
                return;
            }

			Int32 newCode = ToInt32(code);
			string newDescription = ToString(description);
			object[] paramsArray = new object[2];
			paramsArray[0] = newCode;
			paramsArray[1] = newDescription;
			EventBinding.RaiseCustomEvent("OnError", ref paramsArray);
		}

		public void SyncEnd()
        {
            if (!Validate("SyncEnd"))
            {
                return;
            }
           
			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("SyncEnd", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}