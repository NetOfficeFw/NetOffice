using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.ItemEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ItemEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.ItemEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from InspectorsEvents
        /// </summary>
        public static readonly string Id = "0006303A-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public ItemEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}
		
		#endregion

		#region ItemEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
		public void Open([In] [Out] ref object cancel)
        {
            if (!Validate("Open"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Open", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="action"></param>
        /// <param name="response"></param>
        /// <param name="cancel"></param>
        public void CustomAction([In, MarshalAs(UnmanagedType.IDispatch)] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel)
        {
            if (!Validate("CustomAction"))
            {
                Invoker.ReleaseParamsArray(action, response, cancel);
                return;
            }

			object newAction = Factory.CreateEventArgumentObjectFromComProxy(EventClass, action) as object;
			object newResponse = Factory.CreateEventArgumentObjectFromComProxy(EventClass, response) as object;
			object[] paramsArray = new object[3];
			paramsArray[0] = newAction;
			paramsArray[1] = newResponse;
			paramsArray.SetValue(cancel, 2);
			EventBinding.RaiseCustomEvent("CustomAction", ref paramsArray);

			cancel = ToBoolean(paramsArray[2]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        public void CustomPropertyChange([In] object name)
		{
            if (!Validate("CustomPropertyChange"))
            {
                Invoker.ReleaseParamsArray(name);
                return;
            }

			string newName = ToString(name);
			object[] paramsArray = new object[1];
			paramsArray[0] = newName;
			EventBinding.RaiseCustomEvent("CustomPropertyChange", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="forward"></param>
        /// <param name="cancel"></param>
        public void Forward([In, MarshalAs(UnmanagedType.IDispatch)] object forward, [In] [Out] ref object cancel)
        {
            if (!Validate("Forward"))
            {
                Invoker.ReleaseParamsArray(forward, cancel);
                return;
            }

			object newForward = Factory.CreateEventArgumentObjectFromComProxy(EventClass, forward) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newForward;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("Forward", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
        public void Close([In] [Out] ref object cancel)
        {
            if (!Validate("Close"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Close", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        public void PropertyChange([In] object name)
        {
            if (!Validate("PropertyChange"))
            {
                Invoker.ReleaseParamsArray(name);
                return;
            }

			string newName = ToString(name);
			object[] paramsArray = new object[1];
			paramsArray[0] = newName;
			EventBinding.RaiseCustomEvent("PropertyChange", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void Read()
        {
            if (!Validate("Read"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Read", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="response"></param>
        /// <param name="cancel"></param>
        public void Reply([In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel)
        {
            if (!Validate("Reply"))
            {
                Invoker.ReleaseParamsArray(response, cancel);
                return;
            }

			object newResponse = Factory.CreateEventArgumentObjectFromComProxy(EventClass, response) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newResponse;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("Reply", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="response"></param>
        /// <param name="cancel"></param>
        public void ReplyAll([In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel)
		{
            if (!Validate("ReplyAll"))
            {
                Invoker.ReleaseParamsArray(response, cancel);
                return;
            }

            object newResponse = Factory.CreateEventArgumentObjectFromComProxy(EventClass, response) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newResponse;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ReplyAll", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
        public void Send([In] [Out] ref object cancel)
		{
            if (!Validate("Send"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Send", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
        public void Write([In] [Out] ref object cancel)
		{
            if (!Validate("Write"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Write", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
        public void BeforeCheckNames([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeCheckNames"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeCheckNames", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="attachment"></param>
        public void AttachmentAdd([In, MarshalAs(UnmanagedType.IDispatch)] object attachment)
		{
            if (!Validate("AttachmentAdd"))
            {
                Invoker.ReleaseParamsArray(attachment);
                return;
            }

			NetOffice.OutlookApi.Attachment newAttachment = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Attachment>(EventClass, attachment, typeof(NetOffice.OutlookApi.Attachment));
			object[] paramsArray = new object[1];
			paramsArray[0] = newAttachment;
			EventBinding.RaiseCustomEvent("AttachmentAdd", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="attachment"></param>
        public void AttachmentRead([In, MarshalAs(UnmanagedType.IDispatch)] object attachment)
		{
            if (!Validate("AttachmentRead"))
            {
                Invoker.ReleaseParamsArray(attachment);
                return;
            }

            NetOffice.OutlookApi.Attachment newAttachment = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Attachment>(EventClass, attachment, typeof(NetOffice.OutlookApi.Attachment));
            object[] paramsArray = new object[1];
			paramsArray[0] = newAttachment;
			EventBinding.RaiseCustomEvent("AttachmentRead", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="attachment"></param>
        /// <param name="cancel"></param>
        public void BeforeAttachmentSave([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeAttachmentSave"))
            {
                Invoker.ReleaseParamsArray(attachment, cancel);
                return;
            }

            NetOffice.OutlookApi.Attachment newAttachment = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Attachment>(EventClass, attachment, typeof(NetOffice.OutlookApi.Attachment));
            object[] paramsArray = new object[2];
			paramsArray[0] = newAttachment;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeAttachmentSave", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

		#endregion
	}
}

