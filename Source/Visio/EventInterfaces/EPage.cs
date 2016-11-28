using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.VisioApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[ComImport, Guid("000D0B0A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EPage
	{
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8208)]
		void PageChanged([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16400)]
		void BeforePageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32832)]
		void ShapeAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(901)]
		void BeforeSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8256)]
		void ShapeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(902)]
		void SelectionAdded([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16448)]
		void BeforeShapeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8320)]
		void TextChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10240)]
		void CellChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12288)]
		void FormulaChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(33024)]
		void ConnectionsAdded([In, MarshalAs(UnmanagedType.IDispatch)] object connects);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16640)]
		void ConnectionsDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object connects);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(500)]
		void QueryCancelPageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(501)]
		void PageDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(802)]
		void ShapeParentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(803)]
		void BeforeShapeTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(804)]
		void ShapeExitedTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(903)]
		void QueryCancelSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(904)]
		void SelectionDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(905)]
		void QueryCancelUngroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(906)]
		void UngroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(907)]
		void QueryCancelConvertToGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(908)]
		void ConvertToGroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(909)]
		void QueryCancelGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(910)]
		void GroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(807)]
		void ShapeDataGraphicChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(805)]
		void ShapeLinkAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(806)]
		void ShapeLinkDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID);

		[SupportByVersionAttribute("Visio", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(502)]
		void ContainerRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		[SupportByVersionAttribute("Visio", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(503)]
		void ContainerRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		[SupportByVersionAttribute("Visio", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(504)]
		void CalloutRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		[SupportByVersionAttribute("Visio", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(505)]
		void CalloutRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		[SupportByVersionAttribute("Visio", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(911)]
		void QueryCancelReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes);

		[SupportByVersionAttribute("Visio", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(912)]
		void ReplaceShapesCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes);

		[SupportByVersionAttribute("Visio", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(913)]
		void BeforeReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes);

		[SupportByVersionAttribute("Visio", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(914)]
		void AfterReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object sel);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EPage_SinkHelper : SinkHelper, EPage
	{
		#region Static
		
		public static readonly string Id = "000D0B0A-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public EPage_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region EPage Members
		
		public void PageChanged([In, MarshalAs(UnmanagedType.IDispatch)] object page)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PageChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(page);
				return;
			}

			NetOffice.VisioApi.IVPage newPage = Factory.CreateObjectFromComProxy(_eventClass, page) as NetOffice.VisioApi.IVPage;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPage;
			_eventBinding.RaiseCustomEvent("PageChanged", ref paramsArray);
		}

		public void BeforePageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforePageDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(page);
				return;
			}

			NetOffice.VisioApi.IVPage newPage = Factory.CreateObjectFromComProxy(_eventClass, page) as NetOffice.VisioApi.IVPage;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPage;
			_eventBinding.RaiseCustomEvent("BeforePageDelete", ref paramsArray);
		}

		public void ShapeAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ShapeAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shape);
				return;
			}

			NetOffice.VisioApi.IVShape newShape = Factory.CreateObjectFromComProxy(_eventClass, shape) as NetOffice.VisioApi.IVShape;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			_eventBinding.RaiseCustomEvent("ShapeAdded", ref paramsArray);
		}

		public void BeforeSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeSelectionDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection);
				return;
			}

			NetOffice.VisioApi.IVSelection newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as NetOffice.VisioApi.IVSelection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			_eventBinding.RaiseCustomEvent("BeforeSelectionDelete", ref paramsArray);
		}

		public void ShapeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ShapeChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shape);
				return;
			}

			NetOffice.VisioApi.IVShape newShape = Factory.CreateObjectFromComProxy(_eventClass, shape) as NetOffice.VisioApi.IVShape;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			_eventBinding.RaiseCustomEvent("ShapeChanged", ref paramsArray);
		}

		public void SelectionAdded([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SelectionAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection);
				return;
			}

			NetOffice.VisioApi.IVSelection newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as NetOffice.VisioApi.IVSelection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			_eventBinding.RaiseCustomEvent("SelectionAdded", ref paramsArray);
		}

		public void BeforeShapeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeShapeDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shape);
				return;
			}

			NetOffice.VisioApi.IVShape newShape = Factory.CreateObjectFromComProxy(_eventClass, shape) as NetOffice.VisioApi.IVShape;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			_eventBinding.RaiseCustomEvent("BeforeShapeDelete", ref paramsArray);
		}

		public void TextChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("TextChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shape);
				return;
			}

			NetOffice.VisioApi.IVShape newShape = Factory.CreateObjectFromComProxy(_eventClass, shape) as NetOffice.VisioApi.IVShape;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			_eventBinding.RaiseCustomEvent("TextChanged", ref paramsArray);
		}

		public void CellChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CellChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cell);
				return;
			}

			NetOffice.VisioApi.IVCell newCell = Factory.CreateObjectFromComProxy(_eventClass, cell) as NetOffice.VisioApi.IVCell;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCell;
			_eventBinding.RaiseCustomEvent("CellChanged", ref paramsArray);
		}

		public void FormulaChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("FormulaChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cell);
				return;
			}

			NetOffice.VisioApi.IVCell newCell = Factory.CreateObjectFromComProxy(_eventClass, cell) as NetOffice.VisioApi.IVCell;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCell;
			_eventBinding.RaiseCustomEvent("FormulaChanged", ref paramsArray);
		}

		public void ConnectionsAdded([In, MarshalAs(UnmanagedType.IDispatch)] object connects)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ConnectionsAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(connects);
				return;
			}

			NetOffice.VisioApi.IVConnects newConnects = Factory.CreateObjectFromComProxy(_eventClass, connects) as NetOffice.VisioApi.IVConnects;
			object[] paramsArray = new object[1];
			paramsArray[0] = newConnects;
			_eventBinding.RaiseCustomEvent("ConnectionsAdded", ref paramsArray);
		}

		public void ConnectionsDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object connects)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ConnectionsDeleted");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(connects);
				return;
			}

			NetOffice.VisioApi.IVConnects newConnects = Factory.CreateObjectFromComProxy(_eventClass, connects) as NetOffice.VisioApi.IVConnects;
			object[] paramsArray = new object[1];
			paramsArray[0] = newConnects;
			_eventBinding.RaiseCustomEvent("ConnectionsDeleted", ref paramsArray);
		}

		public void QueryCancelPageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelPageDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(page);
				return;
			}

			NetOffice.VisioApi.IVPage newPage = Factory.CreateObjectFromComProxy(_eventClass, page) as NetOffice.VisioApi.IVPage;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPage;
			_eventBinding.RaiseCustomEvent("QueryCancelPageDelete", ref paramsArray);
		}

		public void PageDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object page)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PageDeleteCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(page);
				return;
			}

			NetOffice.VisioApi.IVPage newPage = Factory.CreateObjectFromComProxy(_eventClass, page) as NetOffice.VisioApi.IVPage;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPage;
			_eventBinding.RaiseCustomEvent("PageDeleteCanceled", ref paramsArray);
		}

		public void ShapeParentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ShapeParentChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shape);
				return;
			}

			NetOffice.VisioApi.IVShape newShape = Factory.CreateObjectFromComProxy(_eventClass, shape) as NetOffice.VisioApi.IVShape;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			_eventBinding.RaiseCustomEvent("ShapeParentChanged", ref paramsArray);
		}

		public void BeforeShapeTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeShapeTextEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shape);
				return;
			}

			NetOffice.VisioApi.IVShape newShape = Factory.CreateObjectFromComProxy(_eventClass, shape) as NetOffice.VisioApi.IVShape;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			_eventBinding.RaiseCustomEvent("BeforeShapeTextEdit", ref paramsArray);
		}

		public void ShapeExitedTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ShapeExitedTextEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shape);
				return;
			}

			NetOffice.VisioApi.IVShape newShape = Factory.CreateObjectFromComProxy(_eventClass, shape) as NetOffice.VisioApi.IVShape;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			_eventBinding.RaiseCustomEvent("ShapeExitedTextEdit", ref paramsArray);
		}

		public void QueryCancelSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelSelectionDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection);
				return;
			}

			NetOffice.VisioApi.IVSelection newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as NetOffice.VisioApi.IVSelection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			_eventBinding.RaiseCustomEvent("QueryCancelSelectionDelete", ref paramsArray);
		}

		public void SelectionDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SelectionDeleteCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection);
				return;
			}

			NetOffice.VisioApi.IVSelection newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as NetOffice.VisioApi.IVSelection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			_eventBinding.RaiseCustomEvent("SelectionDeleteCanceled", ref paramsArray);
		}

		public void QueryCancelUngroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelUngroup");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection);
				return;
			}

			NetOffice.VisioApi.IVSelection newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as NetOffice.VisioApi.IVSelection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			_eventBinding.RaiseCustomEvent("QueryCancelUngroup", ref paramsArray);
		}

		public void UngroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("UngroupCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection);
				return;
			}

			NetOffice.VisioApi.IVSelection newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as NetOffice.VisioApi.IVSelection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			_eventBinding.RaiseCustomEvent("UngroupCanceled", ref paramsArray);
		}

		public void QueryCancelConvertToGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelConvertToGroup");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection);
				return;
			}

			NetOffice.VisioApi.IVSelection newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as NetOffice.VisioApi.IVSelection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			_eventBinding.RaiseCustomEvent("QueryCancelConvertToGroup", ref paramsArray);
		}

		public void ConvertToGroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ConvertToGroupCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection);
				return;
			}

			NetOffice.VisioApi.IVSelection newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as NetOffice.VisioApi.IVSelection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			_eventBinding.RaiseCustomEvent("ConvertToGroupCanceled", ref paramsArray);
		}

		public void QueryCancelGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelGroup");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection);
				return;
			}

			NetOffice.VisioApi.IVSelection newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as NetOffice.VisioApi.IVSelection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			_eventBinding.RaiseCustomEvent("QueryCancelGroup", ref paramsArray);
		}

		public void GroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("GroupCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selection);
				return;
			}

			NetOffice.VisioApi.IVSelection newSelection = Factory.CreateObjectFromComProxy(_eventClass, selection) as NetOffice.VisioApi.IVSelection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			_eventBinding.RaiseCustomEvent("GroupCanceled", ref paramsArray);
		}

		public void ShapeDataGraphicChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ShapeDataGraphicChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shape);
				return;
			}

			NetOffice.VisioApi.IVShape newShape = Factory.CreateObjectFromComProxy(_eventClass, shape) as NetOffice.VisioApi.IVShape;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			_eventBinding.RaiseCustomEvent("ShapeDataGraphicChanged", ref paramsArray);
		}

		public void ShapeLinkAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ShapeLinkAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shape, dataRecordsetID, dataRowID);
				return;
			}

			NetOffice.VisioApi.IVShape newShape = Factory.CreateObjectFromComProxy(_eventClass, shape) as NetOffice.VisioApi.IVShape;
			Int32 newDataRecordsetID = Convert.ToInt32(dataRecordsetID);
			Int32 newDataRowID = Convert.ToInt32(dataRowID);
			object[] paramsArray = new object[3];
			paramsArray[0] = newShape;
			paramsArray[1] = newDataRecordsetID;
			paramsArray[2] = newDataRowID;
			_eventBinding.RaiseCustomEvent("ShapeLinkAdded", ref paramsArray);
		}

		public void ShapeLinkDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ShapeLinkDeleted");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shape, dataRecordsetID, dataRowID);
				return;
			}

			NetOffice.VisioApi.IVShape newShape = Factory.CreateObjectFromComProxy(_eventClass, shape) as NetOffice.VisioApi.IVShape;
			Int32 newDataRecordsetID = Convert.ToInt32(dataRecordsetID);
			Int32 newDataRowID = Convert.ToInt32(dataRowID);
			object[] paramsArray = new object[3];
			paramsArray[0] = newShape;
			paramsArray[1] = newDataRecordsetID;
			paramsArray[2] = newDataRowID;
			_eventBinding.RaiseCustomEvent("ShapeLinkDeleted", ref paramsArray);
		}

		public void ContainerRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContainerRelationshipAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shapePair);
				return;
			}

			NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateObjectFromComProxy(_eventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			_eventBinding.RaiseCustomEvent("ContainerRelationshipAdded", ref paramsArray);
		}

		public void ContainerRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ContainerRelationshipDeleted");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shapePair);
				return;
			}

			NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateObjectFromComProxy(_eventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			_eventBinding.RaiseCustomEvent("ContainerRelationshipDeleted", ref paramsArray);
		}

		public void CalloutRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CalloutRelationshipAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shapePair);
				return;
			}

			NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateObjectFromComProxy(_eventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			_eventBinding.RaiseCustomEvent("CalloutRelationshipAdded", ref paramsArray);
		}

		public void CalloutRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CalloutRelationshipDeleted");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shapePair);
				return;
			}

			NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateObjectFromComProxy(_eventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
			object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			_eventBinding.RaiseCustomEvent("CalloutRelationshipDeleted", ref paramsArray);
		}

		public void QueryCancelReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelReplaceShapes");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(replaceShapes);
				return;
			}

			NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateObjectFromComProxy(_eventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
			object[] paramsArray = new object[1];
			paramsArray[0] = newreplaceShapes;
			_eventBinding.RaiseCustomEvent("QueryCancelReplaceShapes", ref paramsArray);
		}

		public void ReplaceShapesCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ReplaceShapesCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(replaceShapes);
				return;
			}

			NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateObjectFromComProxy(_eventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
			object[] paramsArray = new object[1];
			paramsArray[0] = newreplaceShapes;
			_eventBinding.RaiseCustomEvent("ReplaceShapesCanceled", ref paramsArray);
		}

		public void BeforeReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeReplaceShapes");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(replaceShapes);
				return;
			}

			NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateObjectFromComProxy(_eventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
			object[] paramsArray = new object[1];
			paramsArray[0] = newreplaceShapes;
			_eventBinding.RaiseCustomEvent("BeforeReplaceShapes", ref paramsArray);
		}

		public void AfterReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object sel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterReplaceShapes");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sel);
				return;
			}

			NetOffice.VisioApi.IVSelection newsel = Factory.CreateObjectFromComProxy(_eventClass, sel) as NetOffice.VisioApi.IVSelection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newsel;
			_eventBinding.RaiseCustomEvent("AfterReplaceShapes", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}