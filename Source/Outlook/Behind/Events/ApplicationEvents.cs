using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.ApplicationEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.ApplicationEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from ApplicationEvents
        /// </summary>
        public static readonly string Id = "0006304E-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Construction

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public ApplicationEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}

        #endregion

        #region Ctor

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="cancel"></param>
        public void ItemSend([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel)
        {
            if (!Validate("ItemSend"))
            {
                Invoker.ReleaseParamsArray(item, cancel);
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newItem;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ItemSend", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
		}

        /// <summary>
        /// 
        /// </summary>
		public void NewMail()
        {
            if (!Validate("NewMail"))
            { 
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("NewMail", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
		public void Reminder([In, MarshalAs(UnmanagedType.IDispatch)] object item)
        {
            if (!Validate("Reminder"))
            {
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			EventBinding.RaiseCustomEvent("Reminder", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pages"></param>
		public void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages)
        {
            if (!Validate("OptionsPagesAdd"))
            {
                return;
            }

			NetOffice.OutlookApi.PropertyPages newPages = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.PropertyPages>(EventClass, pages, typeof(NetOffice.OutlookApi.PropertyPages));
			object[] paramsArray = new object[1];
			paramsArray[0] = newPages;
			EventBinding.RaiseCustomEvent("OptionsPagesAdd", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void Startup()
        {
            if (!Validate("Startup"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Startup", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void Quit()
        {
            if (!Validate("Quit"))
            {
                return;
            }
            
			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Quit", ref paramsArray);
		}

		#endregion
	}
}

