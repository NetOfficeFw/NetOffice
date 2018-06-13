using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.WordApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.WordApi.EventContracts.DocumentEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class DocumentEvents_SinkHelper : SinkHelper, NetOffice.WordApi.EventContracts.DocumentEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from DocumentEvents
        /// </summary>
        public static readonly string Id = "000209F6-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public DocumentEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region DocumentEvents

        /// <summary>
        /// 
        /// </summary>
        public void New()
        {
            if (!Validate("New"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("New", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        public void Open()
        {
            if (!Validate("Open"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Open", ref paramsArray);
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

        #endregion
    }
}
