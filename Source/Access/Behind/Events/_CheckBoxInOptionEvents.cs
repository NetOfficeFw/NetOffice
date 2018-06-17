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
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts._CheckBoxInOptionEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CheckBoxInOptionEvents_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts._CheckBoxInOptionEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from _CheckBoxInOptionEvents
		/// </summary>
		public static readonly string Id = "BC9E435D-F037-11CD-8701-00AA003F0F07";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public _CheckBoxInOptionEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _CheckBoxInOptionEvents
		
		/// <summary>
		/// 
		/// </summary>
		public void GotFocus()
		{
            if (!Validate("GotFocus"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("GotFocus", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void LostFocus()
		{
            if (!Validate("LostFocus"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("LostFocus", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="shift"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		public void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
        {
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);

			button = ToInt16(paramsArray[0]);
			shift = ToInt16(paramsArray[1]);
			x = ToSingle(paramsArray[2]);
			y = ToSingle(paramsArray[3]);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="shift"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		public void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
        {
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);

            button = ToInt16(paramsArray[0]);
            shift = ToInt16(paramsArray[1]);
            x = ToSingle(paramsArray[2]);
            y = ToSingle(paramsArray[3]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="shift"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		public void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
		{
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);

            button = ToInt16(paramsArray[0]);
            shift = ToInt16(paramsArray[1]);
            x = ToSingle(paramsArray[2]);
            y = ToSingle(paramsArray[3]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyCode"></param>
		/// <param name="shift"></param>
		public void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift)
        {
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray.SetValue(shift, 1);
			EventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);

			keyCode = ToInt16(paramsArray[0]);
			shift = ToSingle(paramsArray[1]);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyAscii"></param>
		public void KeyPress([In] [Out] ref object keyAscii)
        {
            if (!Validate("KeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(keyAscii, 0);
			EventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);

			keyAscii = ToInt16(paramsArray[0]);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyCode"></param>
		/// <param name="shift"></param>
		public void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift)
        {
            if (!Validate("KeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray.SetValue(shift, 1);
			EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);

			keyCode = ToInt16(paramsArray[0]);
			shift = ToInt16(paramsArray[1]);
		}

		#endregion
	}
	
}
