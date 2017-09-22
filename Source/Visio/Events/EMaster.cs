using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Visio", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000D0B08-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EMaster
	{
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8200)]
		void MasterChanged([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16392)]
		void BeforeMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32832)]
		void ShapeAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

        [SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(901)]
		void BeforeSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8256)]
		void ShapeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(902)]
		void SelectionAdded([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16448)]
		void BeforeShapeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8320)]
		void TextChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("cell", typeof(VisioApi.IVCell))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10240)]
		void CellChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("cell", typeof(VisioApi.IVCell))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12288)]
		void FormulaChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("connects", typeof(VisioApi.IVConnects))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(33024)]
		void ConnectionsAdded([In, MarshalAs(UnmanagedType.IDispatch)] object connects);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("connects", typeof(VisioApi.IVConnects))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16640)]
		void ConnectionsDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object connects);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(400)]
		void QueryCancelMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(401)]
		void MasterDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(802)]
		void ShapeParentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(803)]
		void BeforeShapeTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(804)]
		void ShapeExitedTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(903)]
		void QueryCancelSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(904)]
		void SelectionDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(905)]
		void QueryCancelUngroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(906)]
		void UngroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(907)]
		void QueryCancelConvertToGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(908)]
		void ConvertToGroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(909)]
		void QueryCancelGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(910)]
		void GroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(807)]
		void ShapeDataGraphicChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EMaster_SinkHelper : SinkHelper, EMaster
	{
		#region Static
		
		public static readonly string Id = "000D0B08-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public EMaster_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region EMaster
		
        public void MasterChanged([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
            if (!Validate("MasterChanged"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			EventBinding.RaiseCustomEvent("MasterChanged", ref paramsArray);
		}

        public void BeforeMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
            if (!Validate("BeforeMasterDelete"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			EventBinding.RaiseCustomEvent("BeforeMasterDelete", ref paramsArray);
		}

        public void ShapeAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
            if (!Validate("ShapeAdded"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			EventBinding.RaiseCustomEvent("ShapeAdded", ref paramsArray);
		}

        public void BeforeSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("BeforeSelectionDelete"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("BeforeSelectionDelete", ref paramsArray);
		}

        public void ShapeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
            if (!Validate("ShapeChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			EventBinding.RaiseCustomEvent("ShapeChanged", ref paramsArray);
		}

        public void SelectionAdded([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("SelectionAdded"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("SelectionAdded", ref paramsArray);
		}

        public void BeforeShapeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
            if (!Validate("BeforeShapeDelete"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			EventBinding.RaiseCustomEvent("BeforeShapeDelete", ref paramsArray);
		}

        public void TextChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
            if (!Validate("TextChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			EventBinding.RaiseCustomEvent("TextChanged", ref paramsArray);
		}

        public void CellChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell)
		{
            if (!Validate("CellChanged"))
            {
                Invoker.ReleaseParamsArray(cell);
                return;
            }

            NetOffice.VisioApi.IVCell newCell = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cell) as NetOffice.VisioApi.IVCell;
            object[] paramsArray = new object[1];
			paramsArray[0] = newCell;
			EventBinding.RaiseCustomEvent("CellChanged", ref paramsArray);
		}

        public void FormulaChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell)
		{
            if (!Validate("FormulaChanged"))
            {
                Invoker.ReleaseParamsArray(cell);
                return;
            }

            NetOffice.VisioApi.IVCell newCell = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cell) as NetOffice.VisioApi.IVCell;
            object[] paramsArray = new object[1];
			paramsArray[0] = newCell;
			EventBinding.RaiseCustomEvent("FormulaChanged", ref paramsArray);
		}

        public void ConnectionsAdded([In, MarshalAs(UnmanagedType.IDispatch)] object connects)
		{
            if (!Validate("ConnectionsAdded"))
            {
                Invoker.ReleaseParamsArray(connects);
                return;
            }

            NetOffice.VisioApi.IVConnects newConnects = Factory.CreateEventArgumentObjectFromComProxy(EventClass, connects) as NetOffice.VisioApi.IVConnects;
            object[] paramsArray = new object[1];
			paramsArray[0] = newConnects;
			EventBinding.RaiseCustomEvent("ConnectionsAdded", ref paramsArray);
		}

        public void ConnectionsDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object connects)
		{
            if (!Validate("ConnectionsDeleted"))
            {
                Invoker.ReleaseParamsArray(connects);
                return;
            }

            NetOffice.VisioApi.IVConnects newConnects = Factory.CreateEventArgumentObjectFromComProxy(EventClass, connects) as NetOffice.VisioApi.IVConnects;
            object[] paramsArray = new object[1];
			paramsArray[0] = newConnects;
			EventBinding.RaiseCustomEvent("ConnectionsDeleted", ref paramsArray);
		}

        public void QueryCancelMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
            if (!Validate("QueryCancelMasterDelete"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			EventBinding.RaiseCustomEvent("QueryCancelMasterDelete", ref paramsArray);
		}

        public void MasterDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
            if (!Validate("MasterDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			EventBinding.RaiseCustomEvent("MasterDeleteCanceled", ref paramsArray);
		}

        public void ShapeParentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
            if (!Validate("ShapeParentChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			EventBinding.RaiseCustomEvent("ShapeParentChanged", ref paramsArray);
		}

        public void BeforeShapeTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
            if (!Validate("BeforeShapeTextEdit"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			EventBinding.RaiseCustomEvent("BeforeShapeTextEdit", ref paramsArray);
		}

        public void ShapeExitedTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
            if (!Validate("ShapeExitedTextEdit"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			EventBinding.RaiseCustomEvent("ShapeExitedTextEdit", ref paramsArray);
		}

        public void QueryCancelSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("QueryCancelSelectionDelete"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("QueryCancelSelectionDelete", ref paramsArray);
		}

        public void SelectionDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("SelectionDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("SelectionDeleteCanceled", ref paramsArray);
		}

        public void QueryCancelUngroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("QueryCancelUngroup"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("QueryCancelUngroup", ref paramsArray);
		}

        public void UngroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("QueryCancelUngroup"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("UngroupCanceled", ref paramsArray);
		}

        public void QueryCancelConvertToGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("QueryCancelConvertToGroup"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("QueryCancelConvertToGroup", ref paramsArray);
		}

        public void ConvertToGroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("ConvertToGroupCanceled"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("ConvertToGroupCanceled", ref paramsArray);
		}

        public void QueryCancelGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("QueryCancelGroup"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("QueryCancelGroup", ref paramsArray);
		}
        
        public void GroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
		{
            if (!Validate("GroupCanceled"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newSelection;
			EventBinding.RaiseCustomEvent("GroupCanceled", ref paramsArray);
		}
        
        public void ShapeDataGraphicChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
		{
            if (!Validate("ShapeDataGraphicChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShape;
			EventBinding.RaiseCustomEvent("ShapeDataGraphicChanged", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}