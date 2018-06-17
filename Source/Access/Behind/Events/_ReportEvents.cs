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
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts._ReportEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _ReportEvents_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts._ReportEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from _ReportEvents
		/// </summary>
		public static readonly string Id = "BC9E4357-F037-11CD-8701-00AA003F0F07";
		
		#endregion

		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public _ReportEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _ReportEvents
		
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

		#endregion
	}
	
}
