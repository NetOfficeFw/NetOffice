using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.OutlookBarShortcutsEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OutlookBarShortcutsEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.OutlookBarShortcutsEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from OutlookBarShortcutsEvents
        /// </summary>
        public static readonly string Id = "0006307C-0000-0000-C000-000000000046";

        #endregion

        #region Construction

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public OutlookBarShortcutsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}
		
		#endregion

		#region OutlookBarShortcutsEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="newShortcut"></param>
		public void ShortcutAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newShortcut)
		{
            if (!Validate("ShortcutAdd"))
            {
                Invoker.ReleaseParamsArray(newShortcut);
                return;
            }

			NetOffice.OutlookApi.OutlookBarShortcut newNewShortcut = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarShortcut>(EventClass, newShortcut, typeof(NetOffice.OutlookApi.OutlookBarShortcut));
			object[] paramsArray = new object[1];
			paramsArray[0] = newNewShortcut;
			EventBinding.RaiseCustomEvent("ShortcutAdd", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
		public void BeforeShortcutAdd([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeShortcutAdd"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeShortcutAdd", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="shortcut"></param>
        /// <param name="cancel"></param>
		public void BeforeShortcutRemove([In, MarshalAs(UnmanagedType.IDispatch)] object shortcut, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeShortcutRemove"))
            {
                Invoker.ReleaseParamsArray(shortcut, cancel);
                return;
            }

			NetOffice.OutlookApi.OutlookBarShortcut newShortcut = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarShortcut>(EventClass, shortcut, typeof(NetOffice.OutlookApi.OutlookBarShortcut));
			object[] paramsArray = new object[2];
			paramsArray[0] = newShortcut;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeShortcutRemove", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
        }

		#endregion
	}	
}

