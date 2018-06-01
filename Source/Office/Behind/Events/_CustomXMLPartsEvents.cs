using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Exceptions;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OfficeApi.EventContracts._CustomXMLPartsEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _CustomXMLPartsEvents_SinkHelper : SinkHelper, NetOffice.OfficeApi.EventContracts._CustomXMLPartsEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from _CustomXMLPartsEvents
        /// </summary>
        public static readonly string Id = "000CDB0B-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public _CustomXMLPartsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _CustomXMLPartsEvents

        /// <summary>
        /// 
        /// </summary>
        /// <param name="newPart"></param>
        public void PartAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newPart)
        {
            if (!Validate("PartAfterAdd"))
            {
                Invoker.ReleaseParamsArray(newPart);
                return;
            }

            NetOffice.OfficeApi.CustomXMLPart newNewPart = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLPart>(EventClass, newPart, typeof(NetOffice.OfficeApi.CustomXMLPart));
            object[] paramsArray = new object[1];
            paramsArray[0] = newNewPart;
            EventBinding.RaiseCustomEvent("PartAfterAdd", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oldPart"></param>
        public void PartBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldPart)
        {
            if (!Validate("PartBeforeDelete"))
            {
                Invoker.ReleaseParamsArray(oldPart);
                return;
            }

            NetOffice.OfficeApi.CustomXMLPart newOldPart = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLPart>(EventClass, oldPart, typeof(NetOffice.OfficeApi.CustomXMLPart));
            object[] paramsArray = new object[1];
            paramsArray[0] = newOldPart;
            EventBinding.RaiseCustomEvent("PartBeforeDelete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="part"></param>
        public void PartAfterLoad([In, MarshalAs(UnmanagedType.IDispatch)] object part)
        {
            if (!Validate("PartAfterLoad"))
            {
                Invoker.ReleaseParamsArray(part);
                return;
            }

            NetOffice.OfficeApi.CustomXMLPart newPart = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLPart>(EventClass, part, typeof(NetOffice.OfficeApi.CustomXMLPart));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPart;
            EventBinding.RaiseCustomEvent("PartAfterLoad", ref paramsArray);
        }

        #endregion
    }
}
