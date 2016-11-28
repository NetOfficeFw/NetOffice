using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.WordApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Word", 11,12,14,15,16)]
	[ComImport, Guid("00020A02-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DocumentEvents2
	{
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void New();

		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void Open();

		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void Close();

		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void Sync([In] object syncEventType);

		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void XMLAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] object inUndoRedo);

		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void XMLBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object deletedRange, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In] object inUndoRedo);

		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void ContentControlAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newContentControl, [In] object inUndoRedo);

		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void ContentControlBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldContentControl, [In] object inUndoRedo);

		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void ContentControlOnExit([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void ContentControlOnEnter([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl);

		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void ContentControlBeforeStoreUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content);

		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void ContentControlBeforeContentUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content);

		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void BuildingBlockInsert([In, MarshalAs(UnmanagedType.IDispatch)] object range, [In] object name, [In] object category, [In] object blockType, [In] object template);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DocumentEvents2_SinkHelper : SinkHelper, DocumentEvents2
	{
		#region Static
		
		public static readonly string Id = "00020A02-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public DocumentEvents2_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region Properties

        internal Core Factory
        {
            get
            {
                if (null != _eventClass)
                    return _eventClass.Factory;
                else
                    return Core.Default;
            }
        }

        #endregion

		#region DocumentEvents2 Members
		
		public void New()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("New");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("New", ref paramsArray);
		}

		public void Open()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Open");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Open", ref paramsArray);
		}

		public void Close()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Close");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Close", ref paramsArray);
		}

		public void Sync([In] object syncEventType)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Sync");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(syncEventType);
				return;
			}

			NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSyncEventType;
			_eventBinding.RaiseCustomEvent("Sync", ref paramsArray);
		}

		public void XMLAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] object inUndoRedo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("XMLAfterInsert");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(newXMLNode, inUndoRedo);
				return;
			}

			NetOffice.WordApi.XMLNode newNewXMLNode = Factory.CreateObjectFromComProxy(_eventClass, newXMLNode) as NetOffice.WordApi.XMLNode;
			bool newInUndoRedo = Convert.ToBoolean(inUndoRedo);
			object[] paramsArray = new object[2];
			paramsArray[0] = newNewXMLNode;
			paramsArray[1] = newInUndoRedo;
			_eventBinding.RaiseCustomEvent("XMLAfterInsert", ref paramsArray);
		}

		public void XMLBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object deletedRange, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In] object inUndoRedo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("XMLBeforeDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(deletedRange, oldXMLNode, inUndoRedo);
				return;
			}

			NetOffice.WordApi.Range newDeletedRange = Factory.CreateObjectFromComProxy(_eventClass, deletedRange) as NetOffice.WordApi.Range;
			NetOffice.WordApi.XMLNode newOldXMLNode = Factory.CreateObjectFromComProxy(_eventClass, oldXMLNode) as NetOffice.WordApi.XMLNode;
			bool newInUndoRedo = Convert.ToBoolean(inUndoRedo);
			object[] paramsArray = new object[3];
			paramsArray[0] = newDeletedRange;
			paramsArray[1] = newOldXMLNode;
			paramsArray[2] = newInUndoRedo;
			_eventBinding.RaiseCustomEvent("XMLBeforeDelete", ref paramsArray);
		}

		public void ContentControlAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newContentControl, [In] object inUndoRedo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContentControlAfterAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(newContentControl, inUndoRedo);
				return;
			}

			NetOffice.WordApi.ContentControl newNewContentControl = Factory.CreateObjectFromComProxy(_eventClass, newContentControl) as NetOffice.WordApi.ContentControl;
			bool newInUndoRedo = Convert.ToBoolean(inUndoRedo);
			object[] paramsArray = new object[2];
			paramsArray[0] = newNewContentControl;
			paramsArray[1] = newInUndoRedo;
			_eventBinding.RaiseCustomEvent("ContentControlAfterAdd", ref paramsArray);
		}

		public void ContentControlBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldContentControl, [In] object inUndoRedo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContentControlBeforeDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(oldContentControl, inUndoRedo);
				return;
			}

			NetOffice.WordApi.ContentControl newOldContentControl = Factory.CreateObjectFromComProxy(_eventClass, oldContentControl) as NetOffice.WordApi.ContentControl;
			bool newInUndoRedo = Convert.ToBoolean(inUndoRedo);
			object[] paramsArray = new object[2];
			paramsArray[0] = newOldContentControl;
			paramsArray[1] = newInUndoRedo;
			_eventBinding.RaiseCustomEvent("ContentControlBeforeDelete", ref paramsArray);
		}

		public void ContentControlOnExit([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContentControlOnExit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(contentControl, cancel);
				return;
			}

			NetOffice.WordApi.ContentControl newContentControl = Factory.CreateObjectFromComProxy(_eventClass, contentControl) as NetOffice.WordApi.ContentControl;
			object[] paramsArray = new object[2];
			paramsArray[0] = newContentControl;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ContentControlOnExit", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ContentControlOnEnter([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContentControlOnEnter");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(contentControl);
				return;
			}

			NetOffice.WordApi.ContentControl newContentControl = Factory.CreateObjectFromComProxy(_eventClass, contentControl) as NetOffice.WordApi.ContentControl;
			object[] paramsArray = new object[1];
			paramsArray[0] = newContentControl;
			_eventBinding.RaiseCustomEvent("ContentControlOnEnter", ref paramsArray);
		}

		public void ContentControlBeforeStoreUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContentControlBeforeStoreUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(contentControl, content);
				return;
			}

			NetOffice.WordApi.ContentControl newContentControl = Factory.CreateObjectFromComProxy(_eventClass, contentControl) as NetOffice.WordApi.ContentControl;
			object[] paramsArray = new object[2];
			paramsArray[0] = newContentControl;
			paramsArray.SetValue(content, 1);
			_eventBinding.RaiseCustomEvent("ContentControlBeforeStoreUpdate", ref paramsArray);

			content = (string)paramsArray[1];
		}

		public void ContentControlBeforeContentUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContentControlBeforeContentUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(contentControl, content);
				return;
			}

			NetOffice.WordApi.ContentControl newContentControl = Factory.CreateObjectFromComProxy(_eventClass, contentControl) as NetOffice.WordApi.ContentControl;
			object[] paramsArray = new object[2];
			paramsArray[0] = newContentControl;
			paramsArray.SetValue(content, 1);
			_eventBinding.RaiseCustomEvent("ContentControlBeforeContentUpdate", ref paramsArray);

			content = (string)paramsArray[1];
		}

		public void BuildingBlockInsert([In, MarshalAs(UnmanagedType.IDispatch)] object range, [In] object name, [In] object category, [In] object blockType, [In] object template)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BuildingBlockInsert");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(range, name, category, blockType, template);
				return;
			}

			NetOffice.WordApi.Range newRange = Factory.CreateObjectFromComProxy(_eventClass, range) as NetOffice.WordApi.Range;
			string newName = Convert.ToString(name);
			string newCategory = Convert.ToString(category);
			string newBlockType = Convert.ToString(blockType);
			string newTemplate = Convert.ToString(template);
			object[] paramsArray = new object[5];
			paramsArray[0] = newRange;
			paramsArray[1] = newName;
			paramsArray[2] = newCategory;
			paramsArray[3] = newBlockType;
			paramsArray[4] = newTemplate;
			_eventBinding.RaiseCustomEvent("BuildingBlockInsert", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}