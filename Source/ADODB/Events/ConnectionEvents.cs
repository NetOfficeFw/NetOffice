using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi.EventContracts
{
    /// <summary>
    /// ConnectionEvents
    /// </summary>
    [SupportByVersion("ADODB", 2.1,2.5)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
	[ComImport, Guid("00000400-0000-0010-8000-00AA006D2EA4"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ConnectionEvents
	{
        /// <summary>
        /// InfoMessage
        /// </summary>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)]
		void InfoMessage([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

        /// <summary>
        /// BeginTransComplete
        /// </summary>
        /// <param name="transactionLevel"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("transactionLevel", SinkArgumentType.Int32)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void BeginTransComplete([In] object transactionLevel, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

        /// <summary>
        /// CommitTransComplete
        /// </summary>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void CommitTransComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

        /// <summary>
        /// RollbackTransComplete
        /// </summary>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void RollbackTransComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

        /// <summary>
        /// WillExecute
        /// </summary>
        /// <param name="source"></param>
        /// <param name="cursorType"></param>
        /// <param name="lockType"></param>
        /// <param name="options"></param>
        /// <param name="adStatus"></param>
        /// <param name="pCommand"></param>
        /// <param name="pRecordset"></param>
        /// <param name="pConnection"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("source", SinkArgumentType.Enum, typeof(ADODBApi.Enums.CursorTypeEnum))]
        [SinkArgument("cursorType", SinkArgumentType.Enum, typeof(ADODBApi.Enums.LockTypeEnum))]
        [SinkArgument("lockType", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pCommand", typeof(ADODBApi._Command))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void WillExecute([In] [Out] ref object source, [In] object cursorType, [In] object lockType, [In] [Out] ref object options, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pCommand, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

        /// <summary>
        /// ExecuteComplete
        /// </summary>
        /// <param name="recordsAffected"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pCommand"></param>
        /// <param name="pRecordset"></param>
        /// <param name="pConnection"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("recordsAffected", SinkArgumentType.Int32)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pCommand", typeof(ADODBApi._Command))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void ExecuteComplete([In] object recordsAffected, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pCommand, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

        /// <summary>
        /// WillConnect
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="userID"></param>
        /// <param name="password"></param>
        /// <param name="options"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("connectionString", SinkArgumentType.String)]
        [SinkArgument("userID", SinkArgumentType.String)]
        [SinkArgument("password", SinkArgumentType.String)]
        [SinkArgument("options", SinkArgumentType.Int32)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void WillConnect([In] [Out] ref object connectionString, [In] [Out] ref object userID, [In] [Out] ref object password, [In] [Out] ref object options, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

        /// <summary>
        /// ConnectComplete
        /// </summary>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void ConnectComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

        /// <summary>
        /// Disconnect
        /// </summary>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void Disconnect([In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);
	}
}
