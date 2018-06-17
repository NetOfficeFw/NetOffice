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
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts._ObjectFrameEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _ObjectFrameEvents_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts._ObjectFrameEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from _ObjectFrameEvents
		/// </summary>
		public static readonly string Id = "BC9E434D-F037-11CD-8701-00AA003F0F07";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public _ObjectFrameEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion		

		#region _ObjectFrameEvents
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="code"></param>
		public void Updated([In] [Out] ref object code)
		{
            if (!Validate("Updated"))
            {
                Invoker.ReleaseParamsArray(code);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(code, 0);
			EventBinding.RaiseCustomEvent("Updated", ref paramsArray);

			code = ToInt16(paramsArray[0]);
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

			cancel = ToInt16(paramsArray[0]);
        }

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
		/// <param name="cancel"></param>
		public void DblClick([In] [Out] ref object cancel)
        {
            if (!Validate("DblClick"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("DblClick", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
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
            if (!Validate("DblClick"))
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

		#endregion
	}
	
}
