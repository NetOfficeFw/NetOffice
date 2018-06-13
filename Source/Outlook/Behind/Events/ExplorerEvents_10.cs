using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.ExplorerEvents_10"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ExplorerEvents_10_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.ExplorerEvents_10
	{
        #region Static

        /// <summary>
        /// Interface Id from ExplorerEvents_10
        /// </summary>
        public static readonly string Id = "0006300F-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public ExplorerEvents_10_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ExplorerEvents_10
		
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
		public void BeforeMaximize([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeMaximize"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeMaximize", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
		public void BeforeMinimize([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeMinimize"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeMinimize", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
		public void BeforeMove([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeMove"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeMove", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
		public void BeforeSize([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeSize"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeSize", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
		public void BeforeItemCopy([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeItemCopy"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeItemCopy", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
		public void BeforeItemCut([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeItemCut"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeItemCut", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="clipboardContent"></param>
        /// <param name="target"></param>
        /// <param name="cancel"></param>
        public void BeforeItemPaste([In] [Out] ref object clipboardContent, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeItemPaste"))
            {
                Invoker.ReleaseParamsArray(clipboardContent, target, cancel);
                return;
            }
            
            NetOffice.OutlookApi.MAPIFolder newTarget = Factory.CreateEventArgumentObjectFromComProxy(EventClass, target) as NetOffice.OutlookApi.MAPIFolder;
			object[] paramsArray = new object[3];
			paramsArray.SetValue(clipboardContent, 0);
			paramsArray[1] = newTarget;
			paramsArray.SetValue(cancel, 2);
			EventBinding.RaiseCustomEvent("BeforeItemPaste", ref paramsArray);

			clipboardContent = (object)paramsArray[0];
			cancel = ToBoolean(paramsArray[2]);
        }

        /// <summary>
        /// 
        /// </summary>
		public void AttachmentSelectionChange()
        {
            if (!Validate("AttachmentSelectionChange"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AttachmentSelectionChange", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
		public void InlineResponse([In, MarshalAs(UnmanagedType.IDispatch)] object item)
        {
            if (!Validate("InlineResponse"))
            {
                Invoker.ReleaseParamsArray(item);
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			EventBinding.RaiseCustomEvent("InlineResponse", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        public void InlineResponseClose()
        {
            if (!Validate("InlineResponseClose"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("InlineResponseClose", ref paramsArray);
        }

		#endregion
	}	
}
