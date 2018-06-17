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
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts._ReportEvents2"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _ReportEvents2_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts._ReportEvents2
	{
		#region Static
		
		/// <summary>
		/// Interface Id from _ReportEvents2
		/// </summary>
		public static readonly string Id = "D7281A87-4B30-41C5-AB7B-FABF9A35442A";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public _ReportEvents2_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
			
		#region _ReportEvents2
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void Open([In] [Out] ref object cancel)
        {
            if (!Validate("Open"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Open", ref paramsArray);

            cancel = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
		public void Close()
		{
            if (!Validate("Close"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Close", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Activate()
		{
            if (!Validate("Activate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Activate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Deactivate()
		{
            if (!Validate("Deactivate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Deactivate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="dataErr"></param>
		/// <param name="response"></param>
		public void Error([In] [Out] ref object dataErr, [In] [Out] ref object response)
		{
            if (!Validate("Error"))
            {
                Invoker.ReleaseParamsArray(dataErr, response);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(dataErr, 0);
			paramsArray.SetValue(response, 1);
			EventBinding.RaiseCustomEvent("Error", ref paramsArray);

            dataErr = ToInt16(paramsArray[0]);
            response = ToInt16(paramsArray[1]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
        public void NoData([In] [Out] ref object cancel)
        {
            if (!Validate("NoData"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
            paramsArray.SetValue(cancel, 0);
            EventBinding.RaiseCustomEvent("NoData", ref paramsArray);

            cancel = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
        public void Page()
        {
            if (!Validate("Page"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Page", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void Current()
		{
            if (!Validate("Current"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Current", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Load()
		{
            if (!Validate("Load"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Load", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Resize()
		{
            if (!Validate("Resize"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Resize", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void Unload([In] [Out] ref object cancel)
		{
            if (!Validate("Unload"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }
            
			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Unload", ref paramsArray);

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
        public void Timer()
		{
            if (!Validate("Timer"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Timer", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		/// <param name="filterType"></param>
		public void Filter([In] [Out] ref object cancel, [In] [Out] ref object filterType)
        {
            if (!Validate("Filter"))
            {
                Invoker.ReleaseParamsArray(cancel, filterType);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(filterType, 1);
			EventBinding.RaiseCustomEvent("Filter", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
            filterType = ToInt16(paramsArray[1]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		/// <param name="applyType"></param>
		public void ApplyFilter([In] [Out] ref object cancel, [In] [Out] ref object applyType)
        {
            if (!Validate("ApplyFilter"))
            {

                Invoker.ReleaseParamsArray(cancel, applyType);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(applyType, 1);
			EventBinding.RaiseCustomEvent("ApplyFilter", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
            applyType = ToInt16(paramsArray[1]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="page"></param>
		/// <param name="count"></param>
		public void MouseWheel([In] object page, [In] object count)
        {
            if (!Validate("MouseWheel"))
            {
                Invoker.ReleaseParamsArray(page, count);
                return;
            }

            Delegate[] recipients = EventBinding.GetEventRecipients("MouseWheel");
			if( (true == EventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(page, count);
				return;
			}

			bool newPage = ToBoolean(page);
			Int32 newCount = ToInt32(count);
			object[] paramsArray = new object[2];
			paramsArray[0] = newPage;
			paramsArray[1] = newCount;
			EventBinding.RaiseCustomEvent("MouseWheel", ref paramsArray);
		}

		#endregion
	}
	
}
