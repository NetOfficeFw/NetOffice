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
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts.DispPageHdrFtrInReportEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DispPageHdrFtrInReportEvents_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts.DispPageHdrFtrInReportEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from DispPageHdrFtrInReportEvents
		/// </summary>
		public static readonly string Id = "2E705280-92D1-43CC-A57B-ED48BCCC711D";
		
		#endregion	
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public DispPageHdrFtrInReportEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region DispPageHdrFtrInReportEvents Members
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		/// <param name="formatCount"></param>
		public void Format([In] [Out] ref object cancel, [In] [Out] ref object formatCount)
		{
            if (!Validate("Format"))
            {
                Invoker.ReleaseParamsArray(cancel, formatCount);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(formatCount, 1);
			EventBinding.RaiseCustomEvent("Format", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
			formatCount = ToInt16(paramsArray[1]);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		/// <param name="printCount"></param>
		public void Print([In] [Out] ref object cancel, [In] [Out] ref object printCount)
        {
            if (!Validate("Print"))
            {
                Invoker.ReleaseParamsArray(cancel, printCount);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(printCount, 1);
			EventBinding.RaiseCustomEvent("Print", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
            printCount = ToInt16(paramsArray[1]);
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
        public void Paint()
		{
            if (!Validate("Paint"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Paint", ref paramsArray);
		}

		#endregion
	}
	
}
