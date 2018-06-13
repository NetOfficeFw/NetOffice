using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.ExplorerEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ExplorerEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.ExplorerEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from ExplorerEvents
        /// </summary>
        public static readonly string Id = "0006304F-0000-0000-C000-000000000046";
		
		#endregion

		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public ExplorerEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ExplorerEvents
		
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
		public void FolderSwitch()
		{
            if (!Validate("FolderSwitch"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("FolderSwitch", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="newFolder"></param>
        /// <param name="cancel"></param>
		public void BeforeFolderSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object newFolder, [In] [Out] ref object cancel)
		{
            if (!Validate("BeforeFolderSwitch"))
            {
                Invoker.ReleaseParamsArray(newFolder, cancel);
                return;
            }

			object newNewFolder = Factory.CreateEventArgumentObjectFromComProxy(EventClass, newFolder) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newNewFolder;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeFolderSwitch", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        /// <summary>
        /// 
        /// </summary>
		public void ViewSwitch()
		{
            if (!Validate("ViewSwitch"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ViewSwitch", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="newView"></param>
        /// <param name="cancel"></param>
		public void BeforeViewSwitch([In] object newView, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeViewSwitch"))
            {
                Invoker.ReleaseParamsArray(newView, cancel);
                return;
            }

			object newNewView = (object)newView;
			object[] paramsArray = new object[2];
			paramsArray[0] = newNewView;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeViewSwitch", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
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
		public void SelectionChange()
		{
            if (!Validate("SelectionChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("SelectionChange", ref paramsArray);
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
