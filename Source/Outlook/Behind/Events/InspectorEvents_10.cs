using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.InspectorEvents_10"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class InspectorEvents_10_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.InspectorEvents_10
	{
        #region Static

        /// <summary>
        /// Interface Id from InspectorEvents_10
        /// </summary>
        public static readonly string Id = "0006302A-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public InspectorEvents_10_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region InspectorEvents_10
		
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
        /// <param name="activePageName"></param>
		public void PageChange([In] [Out] ref object activePageName)
        {
            if (!Validate("PageChange"))
            {
                Invoker.ReleaseParamsArray(activePageName);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(activePageName, 0);
			EventBinding.RaiseCustomEvent("PageChange", ref paramsArray);

			activePageName = ToString(paramsArray[0]);
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

		#endregion
	}	
}
