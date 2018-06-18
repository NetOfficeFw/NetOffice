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
	/// Default implementation of <see cref="NetOffice.VisioApi.EventContracts.EStyle"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EStyle_SinkHelper : SinkHelper, NetOffice.VisioApi.EventContracts.EStyle
	{
		#region Static
		
		/// <summary>
		/// Interface Id from EStyle
		/// </summary>
		public static readonly string Id = "000D0B06-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public EStyle_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region EStyle

        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        public void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style)
        {
            if (!Validate("StyleChanged"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("StyleChanged", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="style"></param>
		public void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
            if (!Validate("BeforeStyleDelete"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("BeforeStyleDelete", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="style"></param>
		public void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
            if (!Validate("QueryCancelStyleDelete"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("QueryCancelStyleDelete", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="style"></param>
		public void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
            if (!Validate("StyleDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("StyleDeleteCanceled", ref paramsArray);
		}

		#endregion
	}
	
}
