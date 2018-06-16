using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.ADODBApi.EventContracts.RecordsetEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class RecordsetEvents_SinkHelper : SinkHelper, NetOffice.ADODBApi.EventContracts.RecordsetEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from RecordsetEvents
        /// </summary>
        public static readonly string Id = "00000266-0000-0010-8000-00AA006D2EA4";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public RecordsetEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region RecordsetEvents Members

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cFields"></param>
        /// <param name="fields"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
        public virtual void WillChangeField([In] object cFields, [In] object fields, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cFields"></param>
        /// <param name="fields"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
        public virtual void FieldChangeComplete([In] object cFields, [In] object fields, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("FieldChangeComplete"))
            {
                Invoker.ReleaseParamsArray(cFields, fields, pError, adStatus, pRecordset);
                return;
            }

            Int32 newcFields = ToInt32(cFields);
            object newFields = (object)fields;
            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, typeof(NetOffice.ADODBApi.Error));
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="cRecords"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
        public virtual void WillChangeRecord([In] object adReason, [In] object cRecords, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="cRecords"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
        public virtual void RecordChangeComplete([In] object adReason, [In] object cRecords, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("RecordChangeComplete"))
            {
                Invoker.ReleaseParamsArray(adReason, cRecords, pError, adStatus, pRecordset);
                return;
            }

            NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
            Int32 newcRecords = ToInt32(cRecords);
            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, typeof(NetOffice.ADODBApi.Error));
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
        public virtual void WillChangeRecordset([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
        public virtual void RecordsetChangeComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("RecordsetChangeComplete"))
            {
                Invoker.ReleaseParamsArray(adReason, pError, adStatus, pRecordset);
                return;
            }


            NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, typeof(NetOffice.ADODBApi.Error));
            NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[4];
            paramsArray[0] = newadReason;
            paramsArray[1] = newpError;
            paramsArray[2] = newadStatus;
            paramsArray[3] = newpRecordset;
            EventBinding.RaiseCustomEvent("RecordsetChangeComplete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
        public virtual void WillMove([In] object adReason, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="adReason"></param>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
        public virtual void MoveComplete([In] object adReason, [In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("MoveComplete"))
            {
                Invoker.ReleaseParamsArray(adReason, pError, adStatus, pRecordset);
                return;
            }

            NetOffice.ADODBApi.Enums.EventReasonEnum newadReason = (NetOffice.ADODBApi.Enums.EventReasonEnum)adReason;
            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, typeof(NetOffice.ADODBApi.Error));
            NetOffice.ADODBApi.Enums.EventStatusEnum newadStatus = (NetOffice.ADODBApi.Enums.EventStatusEnum)adStatus;
            NetOffice.ADODBApi._Recordset newpRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pRecordset) as NetOffice.ADODBApi._Recordset;
            object[] paramsArray = new object[4];
            paramsArray[0] = newadReason;
            paramsArray[1] = newpError;
            paramsArray[2] = newadStatus;
            paramsArray[3] = newpRecordset;
            EventBinding.RaiseCustomEvent("MoveComplete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fMoreData"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
        public virtual void EndOfRecordset([In] [Out] ref object fMoreData, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="progress"></param>
        /// <param name="maxProgress"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
        public virtual void FetchProgress([In] object progress, [In] object maxProgress, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pError"></param>
        /// <param name="adStatus"></param>
        /// <param name="pRecordset"></param>
        public virtual void FetchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pError, [In] object adStatus, [In, MarshalAs(UnmanagedType.IDispatch)] object pRecordset)
        {
            if (!Validate("FetchComplete"))
            {
                Invoker.ReleaseParamsArray(pError, adStatus, pRecordset);
                return;
            }

            NetOffice.ADODBApi.Error newpError = Factory.CreateKnownObjectFromComProxy<NetOffice.ADODBApi.Error>(EventClass, pError, typeof(NetOffice.ADODBApi.Error));
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
}
