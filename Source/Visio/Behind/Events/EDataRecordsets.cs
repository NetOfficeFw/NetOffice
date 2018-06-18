using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.VisioApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.VisioApi.EventContracts.EDataRecordsets"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EDataRecordsets_SinkHelper : SinkHelper, NetOffice.VisioApi.EventContracts.EDataRecordsets
	{
		#region Static
		
		/// <summary>
		/// Interface Id from EDataRecordsets
		/// </summary>
		public static readonly string Id = "000D0B10-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public EDataRecordsets_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region EDataRecordsets

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataRecordset"></param>
        public void DataRecordsetAdded([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset)
		{
            if (!Validate("DataRecordsetAdded"))
            {
                Invoker.ReleaseParamsArray(dataRecordset);
                return;
            }

            NetOffice.VisioApi.IVDataRecordset newDataRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, dataRecordset) as NetOffice.VisioApi.IVDataRecordset;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordset;
			EventBinding.RaiseCustomEvent("DataRecordsetAdded", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="dataRecordset"></param>
		public void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset)
		{
            if (!Validate("BeforeDataRecordsetDelete"))
            {
                Invoker.ReleaseParamsArray(dataRecordset);
                return;
            }

            NetOffice.VisioApi.IVDataRecordset newDataRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, dataRecordset) as NetOffice.VisioApi.IVDataRecordset;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordset;
			EventBinding.RaiseCustomEvent("BeforeDataRecordsetDelete", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="dataRecordsetChanged"></param>
		public void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged)
		{
            if (!Validate("DataRecordsetChanged"))
            {
                Invoker.ReleaseParamsArray(dataRecordsetChanged);
                return;
            }

            NetOffice.VisioApi.IVDataRecordsetChangedEvent newDataRecordsetChanged = Factory.CreateEventArgumentObjectFromComProxy(EventClass, dataRecordsetChanged) as IVDataRecordsetChangedEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordsetChanged;
			EventBinding.RaiseCustomEvent("DataRecordsetChanged", ref paramsArray);
		}

		#endregion
	}
	
}
