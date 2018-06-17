using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.AccessApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts._References_Events"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _References_Events_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts._References_Events
	{
		#region Static
		
		/// <summary>
		/// Interface Id from _References_Events
		/// </summary>
		public static readonly string Id = "F163F201-ADA2-11CF-89A9-00A0C9054129";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public _References_Events_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}

        #endregion

        #region _References_Events

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

			NetOffice.AccessApi.Reference newReference = Factory.CreateKnownObjectFromComProxy<NetOffice.AccessApi.Reference>(EventClass, reference, typeof(NetOffice.AccessApi.Reference));
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
            if (!Validate("ItemAdded"))
            {
                Invoker.ReleaseParamsArray(reference);
                return;
            }

            NetOffice.AccessApi.Reference newReference = Factory.CreateKnownObjectFromComProxy<NetOffice.AccessApi.Reference>(EventClass, reference, typeof(NetOffice.AccessApi.Reference));
            object[] paramsArray = new object[1];
			paramsArray[0] = newReference;
			EventBinding.RaiseCustomEvent("ItemRemoved", ref paramsArray);
		}

		#endregion
	}
	
}
