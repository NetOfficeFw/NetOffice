using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.VisioApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.VisioApi.EventContracts.ERow"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ERow_SinkHelper : SinkHelper, NetOffice.VisioApi.EventContracts.ERow
	{
		#region Static
		
		/// <summary>
		/// Interface Id from ERow
		/// </summary>
		public static readonly string Id = "000D0B0F-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public ERow_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region ERow

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        public void CellChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell)
		{
            if (!Validate("CellChanged"))
            {
                Invoker.ReleaseParamsArray(cell);
                return;
            }

            NetOffice.VisioApi.IVCell newCell = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cell) as NetOffice.VisioApi.IVCell;
            object[] paramsArray = new object[1];
			paramsArray[0] = newCell;
			EventBinding.RaiseCustomEvent("CellChanged", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cell"></param>
		public void FormulaChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell)
		{
            if (!Validate("FormulaChanged"))
            {
                Invoker.ReleaseParamsArray(cell);
                return;
            }

            NetOffice.VisioApi.IVCell newCell = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cell) as NetOffice.VisioApi.IVCell;
            object[] paramsArray = new object[1];
			paramsArray[0] = newCell;
			EventBinding.RaiseCustomEvent("FormulaChanged", ref paramsArray);
		}

		#endregion
	}
	
}
