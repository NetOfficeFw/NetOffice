using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Access", 14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("EACB9075-68F8-4E3B-B865-E1CE6BE0447C"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DispWebBrowserControlEvents
	{
		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2076)]
		void Updated([In] [Out] ref object code);

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2061)]
		void BeforeUpdate([In] [Out] ref object cancel);

		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2062)]
		void AfterUpdate();

		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2019)]
		void Enter();

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2075)]
		void Exit([In] [Out] ref object cancel);

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2205)]
		void Dirty([In] [Out] ref object cancel);

		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2077)]
		void Change();

		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2073)]
		void GotFocus();

		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2074)]
		void LostFocus();

		[SupportByVersion("Access", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] [Out] ref object cancel);

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

        [SupportByVersion("Access", 14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("keyAscii", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In] [Out] ref object keyAscii);

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("pDisp", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2524)]
		void BeforeNavigate2([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL, [In] [Out] ref object flags, [In] [Out] ref object targetFrameName, [In] [Out] ref object postData, [In] [Out] ref object headers, [In] [Out] ref object cancel);

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("pDisp", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2528)]
		void DocumentComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL);

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("progress", SinkArgumentType.Int32)]
        [SinkArgument("progressMax", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2515)]
		void ProgressChange([In] object progress, [In] object progressMax);

		[SupportByVersion("Access", 14,15,16)]
        [SinkArgument("pDisp", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2510)]
		void NavigateError([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] [Out] ref object uRL, [In] [Out] ref object targetFrameName, [In] [Out] ref object statusCode, [In] [Out] ref object cancel);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DispWebBrowserControlEvents_SinkHelper : SinkHelper, DispWebBrowserControlEvents
	{
		#region Static
		
		public static readonly string Id = "EACB9075-68F8-4E3B-B865-E1CE6BE0447C";
		
		#endregion
		
		#region Ctor

		public DispWebBrowserControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion		

		#region DispWebBrowserControlEvents Members
		
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

        public void AfterUpdate()
        {
            if (!Validate("AfterUpdate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("AfterUpdate", ref paramsArray);
        }

        public void Enter()
        {
            if (!Validate("Enter"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Enter", ref paramsArray);
        }

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

        public void Change()
		{
            if (!Validate("Change"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

        public void GotFocus()
        {
            if (!Validate("GotFocus"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("GotFocus", ref paramsArray);
        }

        public void LostFocus()
        {
            if (!Validate("LostFocus"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("LostFocus", ref paramsArray);
        }

        public void Click()
        {
            if (!Validate("Click"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Click", ref paramsArray);
        }

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
	
	#endregion
	
	#pragma warning restore
}