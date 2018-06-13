using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.OlkLabelEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OlkLabelEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.OlkLabelEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from OlkLabelEvents
        /// </summary>
        public static readonly string Id = "000672E5-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public OlkLabelEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region OlkLabelEvents
		
        /// <summary>
        /// 
        /// </summary>
		public void Click()
		{
            if (!Validate("Click"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Click", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void DoubleClick()
		{
            if (!Validate("DoubleClick"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("DoubleClick", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		public void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y)
		{
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }
           
			NetOffice.OutlookApi.Enums.OlMouseButton newButton = (NetOffice.OutlookApi.Enums.OlMouseButton)button;
			NetOffice.OutlookApi.Enums.OlShiftState newShift = (NetOffice.OutlookApi.Enums.OlShiftState)shift;
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newX;
			paramsArray[3] = newY;
			EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		public void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y)
		{
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }
           
			NetOffice.OutlookApi.Enums.OlMouseButton newButton = (NetOffice.OutlookApi.Enums.OlMouseButton)button;
			NetOffice.OutlookApi.Enums.OlShiftState newShift = (NetOffice.OutlookApi.Enums.OlShiftState)shift;
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newX;
			paramsArray[3] = newY;
			EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		public void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y)
		{
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

            NetOffice.OutlookApi.Enums.OlMouseButton newButton = (NetOffice.OutlookApi.Enums.OlMouseButton)button;
			NetOffice.OutlookApi.Enums.OlShiftState newShift = (NetOffice.OutlookApi.Enums.OlShiftState)shift;
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newX;
			paramsArray[3] = newY;
			EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);
		}

		#endregion
	}	
}
