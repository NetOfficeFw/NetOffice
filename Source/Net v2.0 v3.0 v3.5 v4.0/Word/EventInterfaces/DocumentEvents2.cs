using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.WordApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("Word", 11,12,14)]
	[ComImport, Guid("00020A02-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DocumentEvents2
	{
		[SupportByLibrary("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void New();

		[SupportByLibrary("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void Open();

		[SupportByLibrary("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void Close();

		[SupportByLibrary("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void Sync([In] object syncEventType);

		[SupportByLibrary("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void XMLAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] object inUndoRedo);

		[SupportByLibrary("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void XMLBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object deletedRange, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In] object inUndoRedo);

		[SupportByLibrary("Word", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void ContentControlAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newContentControl, [In] object inUndoRedo);

		[SupportByLibrary("Word", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void ContentControlBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldContentControl, [In] object inUndoRedo);

		[SupportByLibrary("Word", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void ContentControlOnExit([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object cancel);

		[SupportByLibrary("Word", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void ContentControlOnEnter([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl);

		[SupportByLibrary("Word", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void ContentControlBeforeStoreUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content);

		[SupportByLibrary("Word", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void ContentControlBeforeContentUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content);

		[SupportByLibrary("Word", 12,14)]
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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void XMLAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] object inUndoRedo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("XMLAfterInsert");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(newXMLNode, inUndoRedo);
				return;
			}

			NetOffice.WordApi.XMLNode newNewXMLNode = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, newXMLNode) as NetOffice.WordApi.XMLNode;
			bool newInUndoRedo = (bool)inUndoRedo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newNewXMLNode;
			paramsArray[1] = newInUndoRedo;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void XMLBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object deletedRange, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In] object inUndoRedo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("XMLBeforeDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(deletedRange, oldXMLNode, inUndoRedo);
				return;
			}

			NetOffice.WordApi.Range newDeletedRange = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, deletedRange) as NetOffice.WordApi.Range;
			NetOffice.WordApi.XMLNode newOldXMLNode = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, oldXMLNode) as NetOffice.WordApi.XMLNode;
			bool newInUndoRedo = (bool)inUndoRedo;
			object[] paramsArray = new object[3];
			paramsArray[0] = newDeletedRange;
			paramsArray[1] = newOldXMLNode;
			paramsArray[2] = newInUndoRedo;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ContentControlAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newContentControl, [In] object inUndoRedo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContentControlAfterAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(newContentControl, inUndoRedo);
				return;
			}

			NetOffice.WordApi.ContentControl newNewContentControl = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, newContentControl) as NetOffice.WordApi.ContentControl;
			bool newInUndoRedo = (bool)inUndoRedo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newNewContentControl;
			paramsArray[1] = newInUndoRedo;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ContentControlBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldContentControl, [In] object inUndoRedo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContentControlBeforeDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(oldContentControl, inUndoRedo);
				return;
			}

			NetOffice.WordApi.ContentControl newOldContentControl = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, oldContentControl) as NetOffice.WordApi.ContentControl;
			bool newInUndoRedo = (bool)inUndoRedo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newOldContentControl;
			paramsArray[1] = newInUndoRedo;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ContentControlOnExit([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContentControlOnExit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(contentControl, cancel);
				return;
			}

			NetOffice.WordApi.ContentControl newContentControl = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, contentControl) as NetOffice.WordApi.ContentControl;
			object[] paramsArray = new object[2];
			paramsArray[0] = newContentControl;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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

			NetOffice.WordApi.ContentControl newContentControl = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, contentControl) as NetOffice.WordApi.ContentControl;
			object[] paramsArray = new object[1];
			paramsArray[0] = newContentControl;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ContentControlBeforeStoreUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContentControlBeforeStoreUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(contentControl, content);
				return;
			}

			NetOffice.WordApi.ContentControl newContentControl = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, contentControl) as NetOffice.WordApi.ContentControl;
			object[] paramsArray = new object[2];
			paramsArray[0] = newContentControl;
			paramsArray.SetValue(content, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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

			NetOffice.WordApi.ContentControl newContentControl = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, contentControl) as NetOffice.WordApi.ContentControl;
			object[] paramsArray = new object[2];
			paramsArray[0] = newContentControl;
			paramsArray.SetValue(content, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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

			NetOffice.WordApi.Range newRange = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, range) as NetOffice.WordApi.Range;
			string newName = (string)name;
			string newCategory = (string)category;
			string newBlockType = (string)blockType;
			string newTemplate = (string)template;
			object[] paramsArray = new object[5];
			paramsArray[0] = newRange;
			paramsArray[1] = newName;
			paramsArray[2] = newCategory;
			paramsArray[3] = newBlockType;
			paramsArray[4] = newTemplate;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}