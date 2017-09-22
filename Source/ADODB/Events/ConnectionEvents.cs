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
	[ComImport, Guid("00000400-0000-0010-8000-00AA006D2EA4"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ConnectionEvents
	{
		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)]
		void InfoMessage([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("transactionLevel", SinkArgumentType.Int32)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void BeginTransComplete([In] object transactionLevel, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void CommitTransComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void RollbackTransComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("source", SinkArgumentType.Enum, typeof(ADODBApi.Enums.CursorTypeEnum))]
        [SinkArgument("cursorType", SinkArgumentType.Enum, typeof(ADODBApi.Enums.LockTypeEnum))]
        [SinkArgument("lockType", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pCommand", typeof(ADODBApi._Command))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void WillExecute([In] [Out] ref object source, [In] object cursorType, [In] object lockType, [In] [Out] ref object options, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pCommand, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("recordsAffected", SinkArgumentType.Int32)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pCommand", typeof(ADODBApi._Command))]
        [SinkArgument("pRecordset", typeof(ADODBApi._Recordset))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void ExecuteComplete([In] object recordsAffected, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pCommand, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("connectionString", SinkArgumentType.String)]
        [SinkArgument("userID", SinkArgumentType.String)]
        [SinkArgument("password", SinkArgumentType.String)]
        [SinkArgument("options", SinkArgumentType.Int32)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void WillConnect([In] [Out] ref object connectionString, [In] [Out] ref object userID, [In] [Out] ref object password, [In] [Out] ref object options, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("pError", typeof(ADODBApi.Error))]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void ConnectComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);

		[SupportByVersion("ADODB", 2.1,2.5)]
        [SinkArgument("adStatus", SinkArgumentType.Enum, typeof(ADODBApi.Enums.EventStatusEnum))]
        [SinkArgument("pConnection", typeof(ADODBApi._Connection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void Disconnect([In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ConnectionEvents_SinkHelper : SinkHelper, ConnectionEvents
	{
		#region Static
		
		public static readonly string Id = "00000400-0000-0010-8000-00AA006D2EA4";
		
		#endregion	
		
		#region Ctor

		public ConnectionEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region ConnectionEvents Members
		
        public void InfoMessage([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
		{
            if(!Validate("InfoMessage"))
            {
                Invoker.ReleaseParamsArray(pError, adStatus, pConnection);
                return;
            }

			NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, NetOffice.ADODBApi.Error.LateBindingApiWrapperType);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[3];
			paramsArray[0] = newpError;
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpConnection;
			EventBinding.RaiseCustomEvent("InfoMessage", ref paramsArray);
		}

        public void BeginTransComplete([In] object transactionLevel, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (!Validate("BeginTransComplete"))
            {
                Invoker.ReleaseParamsArray(transactionLevel, pError, adStatus, pConnection);
                return;
            }

			Int32 newTransactionLevel = ToInt32(transactionLevel);
			NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, NetOffice.ADODBApi.Error.LateBindingApiWrapperType);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[4];
			paramsArray[0] = newTransactionLevel;
			paramsArray[1] = newpError;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpConnection;
			EventBinding.RaiseCustomEvent("BeginTransComplete", ref paramsArray);
		}

        public void CommitTransComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
		{
            if (!Validate("CommitTransComplete"))
            {
                Invoker.ReleaseParamsArray(pError, adStatus, pConnection);
                return;
            }

			NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, NetOffice.ADODBApi.Error.LateBindingApiWrapperType);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[3];
			paramsArray[0] = newpError;
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpConnection;
			EventBinding.RaiseCustomEvent("CommitTransComplete", ref paramsArray);
		}

        public void RollbackTransComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (!Validate("RollbackTransComplete"))
            {
                Invoker.ReleaseParamsArray(pError, adStatus, pConnection);
                return;
            }

			NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, NetOffice.ADODBApi.Error.LateBindingApiWrapperType);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[3];
			paramsArray[0] = newpError;
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpConnection;
			EventBinding.RaiseCustomEvent("RollbackTransComplete", ref paramsArray);
		}

        public void WillExecute([In] [Out] ref object source, [In] object cursorType, [In] object lockType, [In] [Out] ref object options, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pCommand, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
		{
            if (!Validate("WillExecute"))
            {
                Invoker.ReleaseParamsArray(source, cursorType, lockType, options, adStatus, pCommand, pRecordset, pConnection);
                return;
            }

			NetOffice.ADODBApi.Enums.CursorTypeEnum newCursorType = (NetOffice.ADODBApi.Enums.CursorTypeEnum)cursorType;
			NetOffice.ADODBApi.Enums.LockTypeEnum newLockType = (NetOffice.ADODBApi.Enums.LockTypeEnum)lockType;
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Command newpCommand = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pCommand) as NetOffice.ADODBApi._Command;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[8];
			paramsArray.SetValue(source, 0);
			paramsArray[1] = newCursorType;
			paramsArray[2] = newLockType;
			paramsArray.SetValue(options, 3);
			paramsArray[4] = newadStatus;
			paramsArray[5] = newpCommand;
			paramsArray[6] = newpRecordset;
			paramsArray[7] = newpConnection;
			EventBinding.RaiseCustomEvent("WillExecute", ref paramsArray);

            source = ToString(paramsArray[0]);
            options = ToInt32(paramsArray[3]);            
		}

        public void ExecuteComplete([In] object recordsAffected, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pCommand, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
		{
            if (!Validate("ExecuteComplete"))
            {
                Invoker.ReleaseParamsArray(recordsAffected, pError, adStatus, pCommand, pRecordset, pConnection);
                return;
            }

			Int32 newRecordsAffected = ToInt32(recordsAffected);
			NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, NetOffice.ADODBApi.Error.LateBindingApiWrapperType);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Command newpCommand = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pCommand) as NetOffice.ADODBApi._Command;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[6];
			paramsArray[0] = newRecordsAffected;
			paramsArray[1] = newpError;
			paramsArray[2] = newadStatus;
			paramsArray[3] = newpCommand;
			paramsArray[4] = newpRecordset;
			paramsArray[5] = newpConnection;
			EventBinding.RaiseCustomEvent("ExecuteComplete", ref paramsArray);
		}

        public void WillConnect([In] [Out] ref object connectionString, [In] [Out] ref object userID, [In] [Out] ref object password, [In] [Out] ref object options, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (!Validate("WillConnect"))
            {
                Invoker.ReleaseParamsArray(connectionString, userID, password, options, adStatus, pConnection);
                return;
            }

			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[6];
			paramsArray.SetValue(connectionString, 0);
			paramsArray.SetValue(userID, 1);
			paramsArray.SetValue(password, 2);
			paramsArray.SetValue(options, 3);
			paramsArray[4] = newadStatus;
			paramsArray[5] = newpConnection;
			EventBinding.RaiseCustomEvent("WillConnect", ref paramsArray);

			connectionString = ToString(paramsArray[0]);
			userID = ToString(paramsArray[1]);
			password = ToString(paramsArray[2]);
			options = ToInt32(paramsArray[3]);
		}

        public void ConnectComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (!Validate("ConnectComplete"))
            {
                Invoker.ReleaseParamsArray(pError, adStatus, pConnection);
                return;
            }

			NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, NetOffice.ADODBApi.Error.LateBindingApiWrapperType);
			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[3];
			paramsArray[0] = newpError;
			paramsArray[1] = newadStatus;
			paramsArray[2] = newpConnection;
			EventBinding.RaiseCustomEvent("ConnectComplete", ref paramsArray);
		}

        public void Disconnect([In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (!Validate("Disconnect"))
            {
                Invoker.ReleaseParamsArray(adStatus, pConnection);
                return;
            }

			NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[2];
			paramsArray[0] = newadStatus;
			paramsArray[1] = newpConnection;
			EventBinding.RaiseCustomEvent("Disconnect", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}