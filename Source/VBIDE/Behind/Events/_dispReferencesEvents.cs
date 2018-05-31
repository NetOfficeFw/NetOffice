using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice.Exceptions;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.VBIDEApi.EventContracts._dispReferencesEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _dispReferencesEvents_SinkHelper : SinkHelper, NetOffice.VBIDEApi.EventContracts._dispReferencesEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from _dispReferencesEvents
        /// </summary>
        public static readonly string Id = "0002E118-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public _dispReferencesEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _dispReferencesEvents

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reference"></param>
        public void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
        {
            if (!Validate("ItemAdded"))
            {
                Invoker.ReleaseParamsArray(reference);
                return;
            }

            NetOffice.VBIDEApi.Reference newReference = Factory.CreateKnownObjectFromComProxy<NetOffice.VBIDEApi.Reference>(EventClass, reference, typeof(NetOffice.VBIDEApi.Reference));
            object[] paramsArray = new object[1];
            paramsArray[0] = newReference;
            EventBinding.RaiseCustomEvent("ItemAdded", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reference"></param>
        public void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
        {
            if (!Validate("ItemRemoved"))
            {
                Invoker.ReleaseParamsArray(reference);
                return;
            }

            NetOffice.VBIDEApi.Reference newReference = Factory.CreateKnownObjectFromComProxy<NetOffice.VBIDEApi.Reference>(EventClass, reference, typeof(NetOffice.VBIDEApi.Reference));
            object[] paramsArray = new object[1];
            paramsArray[0] = newReference;
            EventBinding.RaiseCustomEvent("ItemRemoved", ref paramsArray);
        }

        #endregion
    }
}
