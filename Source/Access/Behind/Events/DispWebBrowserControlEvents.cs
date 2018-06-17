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
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts.DispWebBrowserControlEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DispWebBrowserControlEvents_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts.DispWebBrowserControlEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from DispWebBrowserControlEvents
		/// </summary>
		public static readonly string Id = "EACB9075-68F8-4E3B-B865-E1CE6BE0447C";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public DispWebBrowserControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion		

		#region DispWebBrowserControlEvents Members
		
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

            cancel = (Int16)paramsArray[0];
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
		/// <param name="cancel"></param>
        public void Dirty([In] [Out] ref object cancel)
        {
            if (!Validate("Dirty"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
            paramsArray.SetValue(cancel, 0);
            EventBinding.RaiseCustomEvent("Dirty", ref paramsArray);

            cancel = ToInt16(paramsArray[0]);
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
            shift = ToInt16(paramsArray[1]);
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="pDisp"></param>
		/// <param name="uRL"></param>
		/// <param name="flags"></param>
		/// <param name="targetFrameName"></param>
		/// <param name="postData"></param>
		/// <param name="headers"></param>
		/// <param name="cancel"></param>
        public void BeforeNavigate2([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL, [In] [Out] ref object flags, [In] [Out] ref object targetFrameName, [In] [Out] ref object postData, [In] [Out] ref object headers, [In] [Out] ref object cancel)
		{
            if (!Validate("BeforeNavigate2"))
            {
                Invoker.ReleaseParamsArray(pDisp, uRL, flags, targetFrameName, postData, headers, cancel);
                return;
            }

			object newpDisp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pDisp) as object;
			object[] paramsArray = new object[7];
			paramsArray[0] = newpDisp;
			paramsArray.SetValue(uRL, 1);
			paramsArray.SetValue(flags, 2);
			paramsArray.SetValue(targetFrameName, 3);
			paramsArray.SetValue(postData, 4);
			paramsArray.SetValue(headers, 5);
			paramsArray.SetValue(cancel, 6);
			EventBinding.RaiseCustomEvent("BeforeNavigate2", ref paramsArray);

			uRL = (object)paramsArray[1];
			flags = (object)paramsArray[2];
			targetFrameName = (object)paramsArray[3];
			postData = (object)paramsArray[4];
			headers = (object)paramsArray[5];
			cancel = ToBoolean(paramsArray[6]);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="pDisp"></param>
		/// <param name="uRL"></param>
		public void DocumentComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL)
        {
            if (!Validate("DocumentComplete"))
            {
                Invoker.ReleaseParamsArray(pDisp, uRL);
                return;
            }

			object newpDisp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pDisp) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpDisp;
			paramsArray.SetValue(uRL, 1);
			EventBinding.RaiseCustomEvent("DocumentComplete", ref paramsArray);

			uRL = (object)paramsArray[1];
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="progress"></param>
		/// <param name="progressMax"></param>
		public void ProgressChange([In] object progress, [In] object progressMax)
        {
            if (!Validate("ProgressChange"))
            {

                Invoker.ReleaseParamsArray(progress, progressMax);
                return;
            }

			Int32 newProgress = ToInt32(progress);
			Int32 newProgressMax = ToInt32(progressMax);
			object[] paramsArray = new object[2];
			paramsArray[0] = newProgress;
			paramsArray[1] = newProgressMax;
			EventBinding.RaiseCustomEvent("ProgressChange", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="pDisp"></param>
		/// <param name="uRL"></param>
		/// <param name="targetFrameName"></param>
		/// <param name="statusCode"></param>
		/// <param name="cancel"></param>
		public void NavigateError([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL, [In] [Out] ref object targetFrameName, [In] [Out] ref object statusCode, [In] [Out] ref object cancel)
        {
            if (!Validate("NavigateError"))
            {
                Invoker.ReleaseParamsArray(pDisp, uRL, targetFrameName, statusCode, cancel);
                return;
            }

			object newpDisp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, pDisp) as object;
			object[] paramsArray = new object[5];
			paramsArray[0] = newpDisp;
			paramsArray.SetValue(uRL, 1);
			paramsArray.SetValue(targetFrameName, 2);
			paramsArray.SetValue(statusCode, 3);
			paramsArray.SetValue(cancel, 4);
			EventBinding.RaiseCustomEvent("NavigateError", ref paramsArray);

			uRL = (object)paramsArray[1];
			targetFrameName = (object)paramsArray[2];
			statusCode = (object)paramsArray[3];
			cancel = ToBoolean(paramsArray[4]);
		}

		#endregion
	}
	
}
