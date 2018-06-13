using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.OlkCheckBoxEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OlkCheckBoxEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.OlkCheckBoxEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from OlkCheckBoxEvents
        /// </summary>
        public static readonly string Id = "000672E2-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public OlkCheckBoxEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region OlkCheckBoxEvents
		
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

        /// <summary>
        /// 
        /// </summary>
		public void Enter()
		{
            if (!Validate("Enter"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Enter", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
		public void Exit([In] [Out] ref object cancel)
        {
            if (!Validate("Exit"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Exit", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		public void KeyDown([In] [Out] ref object keyCode, [In] object shift)
		{
            if (!Validate("KeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

			NetOffice.OutlookApi.Enums.OlShiftState newShift = (NetOffice.OutlookApi.Enums.OlShiftState)shift;
			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray[1] = newShift;
			EventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);

			keyCode = ToBoolean(paramsArray[0]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="keyAscii"></param>
		public void KeyPress([In] [Out] ref object keyAscii)
        {
            if (!Validate("KeyDown"))
            {
                Invoker.ReleaseParamsArray(keyAscii);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(keyAscii, 0);
			EventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);

			keyAscii = ToInt32(paramsArray[0]);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		public void KeyUp([In] [Out] ref object keyCode, [In] object shift)
        {
            if (!Validate("KeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

			NetOffice.OutlookApi.Enums.OlShiftState newShift = (NetOffice.OutlookApi.Enums.OlShiftState)shift;
			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray[1] = newShift;
			EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);

			keyCode = ToInt32(paramsArray[0]);
		}

        /// <summary>
        /// 
        /// </summary>
		public void Change()
		{
            if (!Validate("Change"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void AfterUpdate()
		{
            if (!Validate("AfterUpdate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AfterUpdate", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
		public void BeforeUpdate([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeUpdate"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeUpdate", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

		#endregion
	}
}
