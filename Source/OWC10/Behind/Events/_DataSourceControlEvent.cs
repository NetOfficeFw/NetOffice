using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OWC10Api.EventContracts._DataSourceControlEvent"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _DataSourceControlEvent_SinkHelper : SinkHelper, NetOffice.OWC10Api.EventContracts._DataSourceControlEvent
    {
        #region Static

        /// <summary>
        /// Interface Id from _DataSourceControlEvent
        /// </summary>
        public static readonly string Id = "F5B39A9B-1480-11D3-8549-00C04FAC67D7";

        #endregion
        
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public _DataSourceControlEvent_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _DataSourceControlEvent

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void Current([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("Current"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("Current", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void BeforeExpand([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforeExpand"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("BeforeExpand", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void BeforeCollapse([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforeCollapse"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("BeforeCollapse", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void BeforeFirstPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforeFirstPage"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("BeforeFirstPage", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void BeforePreviousPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforePreviousPage"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("BeforePreviousPage", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void BeforeNextPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforeNextPage"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("BeforeNextPage", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void BeforeLastPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforeLastPage"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("BeforeLastPage", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void DataError([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("DataError"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("DataError", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void DataPageComplete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("DataPageComplete"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("DataPageComplete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void BeforeInitialBind([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforeInitialBind"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("BeforeInitialBind", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void RecordsetSaveProgress([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("RecordsetSaveProgress"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("RecordsetSaveProgress", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void AfterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("AfterDelete"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("AfterDelete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void AfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("AfterInsert"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("AfterInsert", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void AfterUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("AfterUpdate"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("AfterUpdate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void BeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforeDelete"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("BeforeDelete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void BeforeInsert([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforeInsert"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("BeforeInsert", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void BeforeOverwrite([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforeOverwrite"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("BeforeOverwrite", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void BeforeUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("BeforeUpdate"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("BeforeUpdate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void Dirty([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("Dirty"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("Dirty", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void RecordExit([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("RecordExit"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("RecordExit", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void Undo([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("Undo"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("Undo", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        public void Focus([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo)
        {
            if (!Validate("Focus"))
            {
                Invoker.ReleaseParamsArray(dSCEventInfo);
                return;
            }

            NetOffice.OWC10Api.DSCEventInfo newDSCEventInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.DSCEventInfo>(EventClass, dSCEventInfo, typeof(NetOffice.OWC10Api.DSCEventInfo));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDSCEventInfo;
            EventBinding.RaiseCustomEvent("Focus", ref paramsArray);
        }

        #endregion
    }
}
