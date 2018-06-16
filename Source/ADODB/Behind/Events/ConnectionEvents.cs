using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.ADODBApi.EventContracts.ConnectionEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class ConnectionEvents_SinkHelper : SinkHelper, NetOffice.ADODBApi.EventContracts.ConnectionEvents
    {
        #region Static


        /// <summary>
        /// Interface Id from ConnectionEvents
        /// </summary>
        public static readonly string Id = "00000400-0000-0010-8000-00AA006D2EA4";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public ConnectionEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region ConnectionEvents Members

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
        public virtual void InfoMessage([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (!Validate("InfoMessage"))
            {
                Invoker.ReleaseParamsArray(pError, adStatus, pConnection);
                return;
            }

            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, typeof(NetOffice.ADODBApi.Error));
            NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[3];
            paramsArray[0] = newpError;
            paramsArray[1] = newadStatus;
            paramsArray[2] = newpConnection;
            EventBinding.RaiseCustomEvent("InfoMessage", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="transactionLevel"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
        public virtual void BeginTransComplete([In] object transactionLevel, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (!Validate("BeginTransComplete"))
            {
                Invoker.ReleaseParamsArray(transactionLevel, pError, adStatus, pConnection);
                return;
            }

            Int32 newTransactionLevel = ToInt32(transactionLevel);
            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, typeof(NetOffice.ADODBApi.Error));
            NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[4];
            paramsArray[0] = newTransactionLevel;
            paramsArray[1] = newpError;
            paramsArray[2] = newadStatus;
            paramsArray[3] = newpConnection;
            EventBinding.RaiseCustomEvent("BeginTransComplete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
        public virtual void CommitTransComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (!Validate("CommitTransComplete"))
            {
                Invoker.ReleaseParamsArray(pError, adStatus, pConnection);                return;
            }

            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, typeof(NetOffice.ADODBApi.Error));
            NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[3];
            paramsArray[0] = newpError;
            paramsArray[1] = newadStatus;
            paramsArray[2] = newpConnection;
            EventBinding.RaiseCustomEvent("CommitTransComplete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
        public virtual void RollbackTransComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (!Validate("RollbackTransComplete"))
            {
                Invoker.ReleaseParamsArray(pError, adStatus, pConnection);
                return;
            }

            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, typeof(NetOffice.ADODBApi.Error));
            NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[3];
            paramsArray[0] = newpError;
            paramsArray[1] = newadStatus;
            paramsArray[2] = newpConnection;
            EventBinding.RaiseCustomEvent("RollbackTransComplete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="source"></param>
        /// <param name="cursorType"></param>
        /// <param name="lockType"></param>
        /// <param name="options"></param>
        /// <param name="adStatus"></param>
        /// <param name="pCommand"></param>
        /// <param name="pRecordset"></param>
        /// <param name="pConnection"></param>
        public virtual void WillExecute([In] [Out] ref object source, [In] object cursorType, [In] object lockType, [In] [Out] ref object options, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pCommand, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="recordsAffected"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pCommand"></param>
        /// <param name="pRecordset"></param>
        /// <param name="pConnection"></param>
        public virtual void ExecuteComplete([In] object recordsAffected, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pCommand, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (!Validate("ExecuteComplete"))
            {
                Invoker.ReleaseParamsArray(recordsAffected, pError, adStatus, pCommand, pRecordset, pConnection);
                return;
            }

            Int32 newRecordsAffected = ToInt32(recordsAffected);
            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, typeof(NetOffice.ADODBApi.Error));
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="userID"></param>
        /// <param name="password"></param>
        /// <param name="options"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
        public virtual void WillConnect([In] [Out] ref object connectionString, [In] [Out] ref object userID, [In] [Out] ref object password, [In] [Out] ref object options, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
        public virtual void ConnectComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (!Validate("ConnectComplete"))
            {
                Invoker.ReleaseParamsArray(pError, adStatus, pConnection);
                return;
            }

            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, typeof(NetOffice.ADODBApi.Error));
            NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Connection newpConnection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pConnection) as NetOffice.ADODBApi._Connection;
            object[] paramsArray = new object[3];
            paramsArray[0] = newpError;
            paramsArray[1] = newadStatus;
            paramsArray[2] = newpConnection;
            EventBinding.RaiseCustomEvent("ConnectComplete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="adStatus"></param>
        /// <param name="pConnection"></param>
        public virtual void Disconnect([In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pConnection)
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
}
