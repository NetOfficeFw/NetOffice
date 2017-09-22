using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("ADODB", 2.1,2.5)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00000266-0000-0010-8000-00AA006D2EA4"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface RecordsetEvents
	{
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("cFields", SinkArgumentType.Int32)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void WillChangeField([In] object cFields, [In] object fields, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("cFields", SinkArgumentType.Int32)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void FieldChangeComplete([In] object cFields, [In] object fields, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("cRecords", SinkArgumentType.Int32)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void WillChangeRecord([In] object adReason, [In] object cRecords, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("cRecords", SinkArgumentType.Int32)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void RecordChangeComplete([In] object adReason, [In] object cRecords, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void WillChangeRecordset([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void RecordsetChangeComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void WillMove([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void MoveComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("fMoreData", SinkArgumentType.Bool)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void EndOfRecordset([In] [Out] ref object fMoreData, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("progress", SinkArgumentType.Int32)]
        [SinkArgument("maxProgress", SinkArgumentType.Int32)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void FetchProgress([In] object progress, [In] object maxProgress, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void FetchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class RecordsetEvents_SinkHelper : SinkHelper, RecordsetEvents
	{
		#region Static
		
		public static readonly string Id = "00000266-0000-0010-8000-00AA006D2EA4";
		
		#endregion

		#region Ctor

		public RecordsetEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region RecordsetEvents Members
		
        public void WillChangeField([In] object cFields, [In] object fields, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
            if (!Validate("WillChangeField"))
            {
                Invoker.ReleaseParamsArray(cFields, fields, adStatus, pRecordset);
                return;
            }

			Int32 newcFields = ToInt32(cFields);
			object newFields = (object)fields;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;

            object[] paramsArray = new object[4];
			paramsArray[0] = newcFields;
			paramsArray[1] = newFields;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpRecordset;
			EventBinding.RaiseCustomEvent("WillChangeField", ref paramsArray);
		}

        public void FieldChangeComplete([In] object cFields, [In] object fields, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
            if (!Validate("FieldChangeComplete"))
            {
                Invoker.ReleaseParamsArray(cFields, fields, pError, adStatus, pRecordset);
                return;
            }

			Int32 newcFields = ToInt32(cFields);
			object newFields = (object)fields;
            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, NetOffice.ADODBApi.Error.LateBindingApiWrapperType);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[5];
			paramsArray[0] = newcFields;
			paramsArray[1] = newFields;
			paramsArray[2] = newpError;
			paramsArray[3] = newadStatus;
			paramsArray[4] = newpRecordset;
			EventBinding.RaiseCustomEvent("FieldChangeComplete", ref paramsArray);
		}

        public void WillChangeRecord([In] object adReason, [In] object cRecords, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("WillChangeRecord"))
            {
                Invoker.ReleaseParamsArray(adReason, cRecords, adStatus, pRecordset);
                return;
            }

			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
			Int32 newcRecords = ToInt32(cRecords);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[4];
			paramsArray[0] = newadReason;
			paramsArray[1] = newcRecords;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpRecordset;
			EventBinding.RaiseCustomEvent("WillChangeRecord", ref paramsArray);
		}

        public void RecordChangeComplete([In] object adReason, [In] object cRecords, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("RecordChangeComplete"))
            {
                Invoker.ReleaseParamsArray(adReason, cRecords, pError, adStatus, pRecordset);
                return;
            }

			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
			Int32 newcRecords = ToInt32(cRecords);
			NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, NetOffice.ADODBApi.Error.LateBindingApiWrapperType);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[5];
			paramsArray[0] = newadReason;
			paramsArray[1] = newcRecords;
			paramsArray[2] = newpError;
			paramsArray[3] = newadStatus;
			paramsArray[4] = newpRecordset;
			EventBinding.RaiseCustomEvent("RecordChangeComplete", ref paramsArray);
		}

        public void WillChangeRecordset([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("WillChangeRecordset"))
            {
                Invoker.ReleaseParamsArray(adReason, adStatus, pRecordset);
                return;
            }

			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[3];
			paramsArray[0] = newadReason;
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpRecordset;
			EventBinding.RaiseCustomEvent("WillChangeRecordset", ref paramsArray);
		}

        public void RecordsetChangeComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("RecordsetChangeComplete"))
            {
                Invoker.ReleaseParamsArray(adReason, pError, adStatus, pRecordset);
                return;
            }


			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
			NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, NetOffice.ADODBApi.Error.LateBindingApiWrapperType);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[4];
			paramsArray[0] = newadReason;
			paramsArray[1] = newpError;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpRecordset;
			EventBinding.RaiseCustomEvent("RecordsetChangeComplete", ref paramsArray);
		}

        public void WillMove([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("WillMove"))
            {
                Invoker.ReleaseParamsArray(adReason, adStatus, pRecordset);
                return;
            }

			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[3];
			paramsArray[0] = newadReason;
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpRecordset;
			EventBinding.RaiseCustomEvent("WillMove", ref paramsArray);
		}

        public void MoveComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("MoveComplete"))
            {
                Invoker.ReleaseParamsArray(adReason, pError, adStatus, pRecordset);
                return;
            }

			NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, NetOffice.ADODBApi.Error.LateBindingApiWrapperType);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[4];
			paramsArray[0] = newadReason;
			paramsArray[1] = newpError;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpRecordset;
			EventBinding.RaiseCustomEvent("MoveComplete", ref paramsArray);
		}

        public void EndOfRecordset([In] [Out] ref object fMoreData, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("EndOfRecordset"))
            {
                Invoker.ReleaseParamsArray(fMoreData, adStatus, pRecordset);
                return;
            }

			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[3];
			paramsArray.SetValue(fMoreData, 0);
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpRecordset;
			EventBinding.RaiseCustomEvent("EndOfRecordset", ref paramsArray);

			fMoreData = ToBoolean(paramsArray[0]);
		}

        public void FetchProgress([In] object progress, [In] object maxProgress, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
            if (!Validate("FetchProgress"))
            {
                Invoker.ReleaseParamsArray(progress, maxProgress, adStatus, pRecordset);
                return;
            }

			Int32 newProgress = ToInt32(progress);
			Int32 newMaxProgress = ToInt32(maxProgress);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[4];
			paramsArray[0] = newProgress;
			paramsArray[1] = newMaxProgress;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpRecordset;
			EventBinding.RaiseCustomEvent("FetchProgress", ref paramsArray);
		}

        public void FetchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
		{
            if (!Validate("FetchComplete"))
            {
                Invoker.ReleaseParamsArray(pError, adStatus, pRecordset);
                return;
            }

            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, NetOffice.ADODBApi.Error.LateBindingApiWrapperType);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[3];
			paramsArray[0] = newpError;
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpRecordset;
			EventBinding.RaiseCustomEvent("FetchComplete", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}