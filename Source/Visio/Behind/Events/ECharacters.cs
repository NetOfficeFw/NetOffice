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
	/// Default implementation of <see cref="NetOffice.VisioApi.EventContracts.ECharacters"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ECharacters_SinkHelper : SinkHelper, NetOffice.VisioApi.EventContracts.ECharacters
	{
		#region Static
		
		/// <summary>
		/// Interface Id from ECharacters
		/// </summary>
		public static readonly string Id = "000D0B0C-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public ECharacters_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region ECharacters

        /// <summary>
        /// 
        /// </summary>
        /// <param name="shape"></param>
        public void TextChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("TextChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			EventBinding.RaiseCustomEvent("TextChanged", ref paramsArray);
		}

		#endregion
	}
	
}
