using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.PowerPointApi.EventContracts.OCXExtenderEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class OCXExtenderEvents_SinkHelper : SinkHelper, NetOffice.PowerPointApi.EventContracts.OCXExtenderEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from OCXExtenderEvents
        /// </summary>
        public static readonly string Id = "914934C1-5A91-11CF-8700-00AA0060263B";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public OCXExtenderEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region OCXExtenderEvents

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
