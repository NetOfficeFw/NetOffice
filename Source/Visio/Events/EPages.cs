using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.EventContracts
{
    /// <summary>
    /// EPages
    /// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000D0B09-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EPages
	{
		/// <summary>
		/// PageAdded
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("page", typeof(VisioApi.IVPage))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32784)]
		void PageAdded([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		/// <summary>
		/// PageChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("page", typeof(VisioApi.IVPage))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8208)]
		void PageChanged([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		/// <summary>
		/// BeforePageDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("page", typeof(VisioApi.IVPage))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16400)]
		void BeforePageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		/// <summary>
		/// ShapeAdded
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32832)]
		void ShapeAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		/// <summary>
		/// BeforeSelectionDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(901)]
		void BeforeSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		/// <summary>
		/// ShapeChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8256)]
		void ShapeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		/// <summary>
		/// SelectionAdded
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(902)]
		void SelectionAdded([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		/// <summary>
		/// BeforeShapeDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16448)]
		void BeforeShapeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		/// <summary>
		/// TextChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8320)]
		void TextChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		/// <summary>
		/// CellChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("cell", typeof(VisioApi.IVCell))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10240)]
		void CellChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell);

		/// <summary>
		/// FormulaChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("cell", typeof(VisioApi.IVCell))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12288)]
		void FormulaChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell);

		/// <summary>
		/// ConnectionsAdded
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("connects", typeof(VisioApi.IVConnects))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(33024)]
		void ConnectionsAdded([In, MarshalAs(UnmanagedType.IDispatch)] object connects);

		/// <summary>
		/// ConnectionsDeleted
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("connects", typeof(VisioApi.IVConnects))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16640)]
		void ConnectionsDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object connects);

		/// <summary>
		/// QueryCancelPageDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("page", typeof(VisioApi.IVPage))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(500)]
		void QueryCancelPageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		/// <summary>
		/// PageDeleteCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("page", typeof(VisioApi.IVPage))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(501)]
		void PageDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		/// <summary>
		/// ShapeParentChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(802)]
		void ShapeParentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		/// <summary>
		/// BeforeShapeTextEdit
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(803)]
		void BeforeShapeTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		/// <summary>
		/// ShapeExitedTextEdit
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(804)]
		void ShapeExitedTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		/// <summary>
		/// QueryCancelSelectionDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(903)]
		void QueryCancelSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		/// <summary>
		/// SelectionDeleteCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(904)]
		void SelectionDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		/// <summary>
		/// QueryCancelUngroup
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(905)]
		void QueryCancelUngroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		/// <summary>
		/// UngroupCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(906)]
		void UngroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		/// <summary>
		/// QueryCancelConvertToGroup
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(907)]
		void QueryCancelConvertToGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		/// <summary>
		/// ConvertToGroupCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(908)]
		void ConvertToGroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		/// <summary>
		/// QueryCancelGroup
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(909)]
		void QueryCancelGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		/// <summary>
		/// GroupCanceled
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("selection", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(910)]
		void GroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection);

		/// <summary>
		/// ShapeDataGraphicChanged
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(807)]
		void ShapeDataGraphicChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		/// <summary>
		/// ShapeLinkAdded
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [SinkArgument("dataRecordsetID", SinkArgumentType.Int32)]
        [SinkArgument("dataRowID", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(805)]
		void ShapeLinkAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID);

		/// <summary>
		/// ShapeLinkDeleted
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [SinkArgument("dataRecordsetID", SinkArgumentType.Int32)]
        [SinkArgument("dataRowID", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(806)]
		void ShapeLinkDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID);

		/// <summary>
		/// ContainerRelationshipAdded
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
        [SinkArgument("shapePair", typeof(VisioApi.IVRelatedShapePairEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(502)]
		void ContainerRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		/// <summary>
		/// ContainerRelationshipDeleted
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
        [SinkArgument("shapePair", typeof(VisioApi.IVRelatedShapePairEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(503)]
		void ContainerRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		/// <summary>
		/// CalloutRelationshipAdded
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
        [SinkArgument("shapePair", typeof(VisioApi.IVRelatedShapePairEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(504)]
		void CalloutRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		/// <summary>
		/// CalloutRelationshipDeleted
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
        [SinkArgument("shapePair", typeof(VisioApi.IVRelatedShapePairEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(505)]
		void CalloutRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		/// <summary>
		/// QueryCancelReplaceShapes
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
        [SinkArgument("replaceShapes", typeof(VisioApi.IVReplaceShapesEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(911)]
		void QueryCancelReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes);

		/// <summary>
		/// ReplaceShapesCanceled
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
        [SinkArgument("replaceShapes", typeof(VisioApi.IVReplaceShapesEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(912)]
		void ReplaceShapesCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes);

		/// <summary>
		/// BeforeReplaceShapes
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
        [SinkArgument("replaceShapes", typeof(VisioApi.IVReplaceShapesEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(913)]
		void BeforeReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes);

		/// <summary>
		/// AfterReplaceShapes
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
        [SinkArgument("sel", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(914)]
		void AfterReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object sel);
	}

}
