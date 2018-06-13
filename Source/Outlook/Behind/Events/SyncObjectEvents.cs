using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.SyncObjectEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class SyncObjectEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.SyncObjectEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from SyncObjectEvents
        /// </summary>
        public static readonly string Id = "00063085-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public SyncObjectEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region SyncObjectEvents
		
        /// <summary>
        /// 
        /// </summary>
		public void SyncStart()
		{
            if (!Validate("SyncStart"))
            {              
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("SyncStart", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="state"></param>
        /// <param name="description"></param>
        /// <param name="value"></param>
        /// <param name="max"></param>
        public void Progress([In] object state, [In] object description, [In] object value, [In] object max)
        {
            if (!Validate("Progress"))
            {
                Invoker.ReleaseParamsArray(state, description, value, max);
                return;
            }

			NetOffice.OutlookApi.Enums.OlSyncState newState = (NetOffice.OutlookApi.Enums.OlSyncState)state;
			string newDescription = ToString(description);
			Int32 newValue = ToInt32(value);
			Int32 newMax = ToInt32(max);
			object[] paramsArray = new object[4];
			paramsArray[0] = newState;
			paramsArray[1] = newDescription;
			paramsArray[2] = newValue;
			paramsArray[3] = newMax;
			EventBinding.RaiseCustomEvent("Progress", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="code"></param>
        /// <param name="description"></param>
		public void OnError([In] object code, [In] object description)
        {
            if (!Validate("OnError"))
            {
                Invoker.ReleaseParamsArray(code, description);
                return;
            }

			Int32 newCode = ToInt32(code);
			string newDescription = ToString(description);
			object[] paramsArray = new object[2];
			paramsArray[0] = newCode;
			paramsArray[1] = newDescription;
			EventBinding.RaiseCustomEvent("OnError", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void SyncEnd()
        {
            if (!Validate("SyncEnd"))
            {
                return;
            }
           
			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("SyncEnd", ref paramsArray);
		}

		#endregion
	}	
}
