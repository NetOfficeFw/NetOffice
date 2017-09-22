using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.WordApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Word", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00020A02-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DocumentEvents2
	{
		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void New();

		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void Open();

		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void Close();

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("syncEventType", SinkArgumentType.Enum, typeof(OfficeApi.Enums.MsoSyncEventType))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void Sync([In] object syncEventType);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("newXMLNode", typeof(WordApi.XMLNode))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void XMLAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] object inUndoRedo);

		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("deletedRange", typeof(WordApi.Range))]
        [SinkArgument("oldXMLNode", typeof(WordApi.XMLNode))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void XMLBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object deletedRange, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In] object inUndoRedo);

		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("newContentControl", typeof(WordApi.ContentControl))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void ContentControlAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newContentControl, [In] object inUndoRedo);

		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("oldContentControl", typeof(WordApi.ContentControl))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void ContentControlBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldContentControl, [In] object inUndoRedo);

		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("contentControl", typeof(WordApi.ContentControl))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void ContentControlOnExit([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object cancel);

		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("contentControl", typeof(WordApi.ContentControl))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void ContentControlOnEnter([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl);

		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("contentControl", typeof(WordApi.ContentControl))]
        [SinkArgument("content", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void ContentControlBeforeStoreUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content);

		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("contentControl", typeof(WordApi.ContentControl))]
        [SinkArgument("content", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void ContentControlBeforeContentUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content);

		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("range", typeof(WordApi.Range))]
        [SinkArgument("name", SinkArgumentType.String)]
        [SinkArgument("category", SinkArgumentType.String)]
        [SinkArgument("blockType", SinkArgumentType.String)]
        [SinkArgument("template", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void BuildingBlockInsert([In, MarshalAs(UnmanagedType.IDispatch)] object range, [In] object name, [In] object category, [In] object blockType, [In] object template);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DocumentEvents2_SinkHelper : SinkHelper, DocumentEvents2
	{
		#region Static
		
		public static readonly string Id = "00020A02-0000-0000-C000-000000000046";
		
		#endregion
		
		#region Ctor

		public DocumentEvents2_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}

        #endregion

        #region DocumentEvents2

        public void New()
        {
            if (!Validate("New"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("New", ref paramsArray);
        }

        public void Open()
        {
            if (!Validate("Open"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Open", ref paramsArray);
        }

        public void Close()
        {
            if (!Validate("Close"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Close", ref paramsArray);
        }

        public void Sync([In] object syncEventType)
		{
            if (!Validate("Sync"))
            {
                Invoker.ReleaseParamsArray(syncEventType);
                return;
            }

			NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSyncEventType;
			EventBinding.RaiseCustomEvent("Sync", ref paramsArray);
		}
 
        public void XMLAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] object inUndoRedo)
        {
            if (!Validate("XMLAfterInsert"))
            {
                Invoker.ReleaseParamsArray(newXMLNode, inUndoRedo);
                return;
            }

			NetOffice.WordApi.XMLNode newNewXMLNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.XMLNode>(EventClass, newXMLNode, NetOffice.WordApi.XMLNode.LateBindingApiWrapperType);
			bool newInUndoRedo = ToBoolean(inUndoRedo);
			object[] paramsArray = new object[2];
			paramsArray[0] = newNewXMLNode;
			paramsArray[1] = newInUndoRedo;
			EventBinding.RaiseCustomEvent("XMLAfterInsert", ref paramsArray);
		}

        public void XMLBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object deletedRange, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In] object inUndoRedo)
        {
            if (!Validate("XMLBeforeDelete"))
            {
                Invoker.ReleaseParamsArray(deletedRange, oldXMLNode, inUndoRedo);
                return;
            }

			NetOffice.WordApi.Range newDeletedRange = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Range>(EventClass, deletedRange, NetOffice.WordApi.Range.LateBindingApiWrapperType);
			NetOffice.WordApi.XMLNode newOldXMLNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.XMLNode>(EventClass, oldXMLNode, NetOffice.WordApi.XMLNode.LateBindingApiWrapperType);
			bool newInUndoRedo = Convert.ToBoolean(inUndoRedo);
			object[] paramsArray = new object[3];
			paramsArray[0] = newDeletedRange;
			paramsArray[1] = newOldXMLNode;
			paramsArray[2] = newInUndoRedo;
			EventBinding.RaiseCustomEvent("XMLBeforeDelete", ref paramsArray);
		}

        public void ContentControlAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newContentControl, [In] object inUndoRedo)
        {
            if (!Validate("ContentControlAfterAdd"))
            {
                Invoker.ReleaseParamsArray(newContentControl, inUndoRedo);
                return;
            }

			NetOffice.WordApi.ContentControl newNewContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, newContentControl, NetOffice.WordApi.ContentControl.LateBindingApiWrapperType);
			bool newInUndoRedo = ToBoolean(inUndoRedo);
			object[] paramsArray = new object[2];
			paramsArray[0] = newNewContentControl;
			paramsArray[1] = newInUndoRedo;
			EventBinding.RaiseCustomEvent("ContentControlAfterAdd", ref paramsArray);
		}

        public void ContentControlBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldContentControl, [In] object inUndoRedo)
		{
            if (!Validate("ContentControlBeforeDelete"))
            {
                Invoker.ReleaseParamsArray(oldContentControl, inUndoRedo);
                return;
            }

			NetOffice.WordApi.ContentControl newOldContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, oldContentControl, NetOffice.WordApi.ContentControl.LateBindingApiWrapperType);
            bool newInUndoRedo = Convert.ToBoolean(inUndoRedo);
			object[] paramsArray = new object[2];
			paramsArray[0] = newOldContentControl;
			paramsArray[1] = newInUndoRedo;
			EventBinding.RaiseCustomEvent("ContentControlBeforeDelete", ref paramsArray);
		}

        public void ContentControlOnExit([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object cancel)
		{
            if (!Validate("ContentControlOnExit"))
            {
                Invoker.ReleaseParamsArray(contentControl, cancel);
                return;
            }

			NetOffice.WordApi.ContentControl newContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, contentControl, NetOffice.WordApi.ContentControl.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newContentControl;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ContentControlOnExit", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        public void ContentControlOnEnter([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl)
        {
            if (!Validate("ContentControlOnEnter"))
            {
                Invoker.ReleaseParamsArray(contentControl);
                return;
            }

			NetOffice.WordApi.ContentControl newContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, contentControl, NetOffice.WordApi.ContentControl.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newContentControl;
			EventBinding.RaiseCustomEvent("ContentControlOnEnter", ref paramsArray);
		}

        public void ContentControlBeforeStoreUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content)
        {
            if (!Validate("ContentControlBeforeStoreUpdate"))
            {
                Invoker.ReleaseParamsArray(contentControl, content);
                return;
            }

            NetOffice.WordApi.ContentControl newContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, contentControl, NetOffice.WordApi.ContentControl.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newContentControl;
			paramsArray.SetValue(content, 1);
			EventBinding.RaiseCustomEvent("ContentControlBeforeStoreUpdate", ref paramsArray);

			content = ToString(paramsArray[1]);
		}

        public void ContentControlBeforeContentUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content)
        {
            if (!Validate("ContentControlBeforeStoreUpdate"))
            {
                Invoker.ReleaseParamsArray(contentControl, content);
                return;
            }

			NetOffice.WordApi.ContentControl newContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, contentControl, NetOffice.WordApi.ContentControl.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newContentControl;
			paramsArray.SetValue(content, 1);
			EventBinding.RaiseCustomEvent("ContentControlBeforeContentUpdate", ref paramsArray);

            content = ToString(paramsArray[1]);
        }

        public void BuildingBlockInsert([In, MarshalAs(UnmanagedType.IDispatch)] object range, [In] object name, [In] object category, [In] object blockType, [In] object template)
        {
            if (!Validate("BuildingBlockInsert"))
            {
                Invoker.ReleaseParamsArray(range, name, category, blockType, template);
                return;
            }

			NetOffice.WordApi.Range newRange = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Range>(EventClass, range, NetOffice.WordApi.Range.LateBindingApiWrapperType);
			string newName = ToString(name);
			string newCategory = ToString(category);
			string newBlockType = ToString(blockType);
			string newTemplate = ToString(template);
			object[] paramsArray = new object[5];
			paramsArray[0] = newRange;
			paramsArray[1] = newName;
			paramsArray[2] = newCategory;
			paramsArray[3] = newBlockType;
			paramsArray[4] = newTemplate;
			EventBinding.RaiseCustomEvent("BuildingBlockInsert", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}