using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.OfficeApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[ComImport, Guid("000CDB07-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomXMLPartEvents
	{
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void NodeAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newNode, [In] object inUndoRedo);

		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void NodeAfterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldNode, [In, MarshalAs(UnmanagedType.IDispatch)] object oldParentNode, [In, MarshalAs(UnmanagedType.IDispatch)] object oldNextSibling, [In] object inUndoRedo);

		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void NodeAfterReplace([In, MarshalAs(UnmanagedType.IDispatch)] object oldNode, [In, MarshalAs(UnmanagedType.IDispatch)] object newNode, [In] object inUndoRedo);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CustomXMLPartEvents_SinkHelper : SinkHelper, _CustomXMLPartEvents
	{
		#region Static
		
		public static readonly string Id = "000CDB07-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _CustomXMLPartEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region _CustomXMLPartEvents Members
		
		public void NodeAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newNode, [In] object inUndoRedo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NodeAfterInsert");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(newNode, inUndoRedo);
				return;
			}

			NetOffice.OfficeApi.CustomXMLNode newNewNode = Factory.CreateObjectFromComProxy(_eventClass, newNode) as NetOffice.OfficeApi.CustomXMLNode;
			bool newInUndoRedo = Convert.ToBoolean(inUndoRedo);
			object[] paramsArray = new object[2];
			paramsArray[0] = newNewNode;
			paramsArray[1] = newInUndoRedo;
			_eventBinding.RaiseCustomEvent("NodeAfterInsert", ref paramsArray);
		}

		public void NodeAfterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldNode, [In, MarshalAs(UnmanagedType.IDispatch)] object oldParentNode, [In, MarshalAs(UnmanagedType.IDispatch)] object oldNextSibling, [In] object inUndoRedo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NodeAfterDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(oldNode, oldParentNode, oldNextSibling, inUndoRedo);
				return;
			}

			NetOffice.OfficeApi.CustomXMLNode newOldNode = Factory.CreateObjectFromComProxy(_eventClass, oldNode) as NetOffice.OfficeApi.CustomXMLNode;
			NetOffice.OfficeApi.CustomXMLNode newOldParentNode = Factory.CreateObjectFromComProxy(_eventClass, oldParentNode) as NetOffice.OfficeApi.CustomXMLNode;
			NetOffice.OfficeApi.CustomXMLNode newOldNextSibling = Factory.CreateObjectFromComProxy(_eventClass, oldNextSibling) as NetOffice.OfficeApi.CustomXMLNode;
			bool newInUndoRedo = Convert.ToBoolean(inUndoRedo);
			object[] paramsArray = new object[4];
			paramsArray[0] = newOldNode;
			paramsArray[1] = newOldParentNode;
			paramsArray[2] = newOldNextSibling;
			paramsArray[3] = newInUndoRedo;
			_eventBinding.RaiseCustomEvent("NodeAfterDelete", ref paramsArray);
		}

		public void NodeAfterReplace([In, MarshalAs(UnmanagedType.IDispatch)] object oldNode, [In, MarshalAs(UnmanagedType.IDispatch)] object newNode, [In] object inUndoRedo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NodeAfterReplace");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(oldNode, newNode, inUndoRedo);
				return;
			}

			NetOffice.OfficeApi.CustomXMLNode newOldNode = Factory.CreateObjectFromComProxy(_eventClass, oldNode) as NetOffice.OfficeApi.CustomXMLNode;
			NetOffice.OfficeApi.CustomXMLNode newNewNode = Factory.CreateObjectFromComProxy(_eventClass, newNode) as NetOffice.OfficeApi.CustomXMLNode;
			bool newInUndoRedo = Convert.ToBoolean(inUndoRedo);
			object[] paramsArray = new object[3];
			paramsArray[0] = newOldNode;
			paramsArray[1] = newNewNode;
			paramsArray[2] = newInUndoRedo;
			_eventBinding.RaiseCustomEvent("NodeAfterReplace", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}