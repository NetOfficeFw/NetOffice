using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi.EventContracts
{
    /// <summary>
    /// RecordsetEvents
    /// </summary>
    [SupportByVersion("ADODB", 2.1,2.5)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00000266-0000-0010-8000-00AA006D2EA4"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface RecordsetEvents
	{
        /// <summary>
        /// WillChangeField
        /// </summary>
        /// <param name="cFields"></param>
        /// <param name="fields"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("cFields", SinkArgumentType.Int32)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void WillChangeField([In] object cFields, [In] object fields, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

        /// <summary>
        /// FieldChangeComplete
        /// </summary>
        /// <param name="cFields"></param>
        /// <param name="fields"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("cFields", SinkArgumentType.Int32)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void FieldChangeComplete([In] object cFields, [In] object fields, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

        /// <summary>
        /// WillChangeRecord
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="cRecords"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("cRecords", SinkArgumentType.Int32)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void WillChangeRecord([In] object adReason, [In] object cRecords, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

        /// <summary>
        /// RecordChangeComplete
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="cRecords"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("cRecords", SinkArgumentType.Int32)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void RecordChangeComplete([In] object adReason, [In] object cRecords, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

        /// <summary>
        /// WillChangeRecordset
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void WillChangeRecordset([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

        /// <summary>
        /// RecordsetChangeComplete
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void RecordsetChangeComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

        /// <summary>
        /// WillMove
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void WillMove([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

        /// <summary>
        /// MoveComplete
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adReason", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventReasonEnum))]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void MoveComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

        /// <summary>
        /// EndOfRecordset
        /// </summary>
        /// <param name="fMoreData"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("fMoreData", SinkArgumentType.Bool)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void EndOfRecordset([In] [Out] ref object fMoreData, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

        /// <summary>
        /// FetchProgress
        /// </summary>
        /// <param name="progress"></param>
        /// <param name="maxProgress"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("progress", SinkArgumentType.Int32)]
        [SinkArgument("maxProgress", SinkArgumentType.Int32)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void FetchProgress([In] object progress, [In] object maxProgress, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);

        /// <summary>
        /// FetchComplete
        /// </summary>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void FetchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset);
	}
}
