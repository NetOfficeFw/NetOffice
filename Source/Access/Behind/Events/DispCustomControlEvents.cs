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
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts.DispCustomControlEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DispCustomControlEvents_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts.DispCustomControlEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from DispCustomControlEvents
		/// </summary>
		public static readonly string Id = "2E70526E-92D1-43CC-A57B-ED48BCCC711D";
		
		#endregion

		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public DispCustomControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region DispCustomControlEvents
		
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

        #endregion
    }
	
}
