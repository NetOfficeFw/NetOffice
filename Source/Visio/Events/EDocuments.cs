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
    [ComImport, Guid("000D0B03-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EDocuments
	{
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void DocumentOpened([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void DocumentCreated([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void DocumentSaved([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void DocumentSavedAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8194)]
		void DocumentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16386)]
		void BeforeDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32772)]
		void StyleAdded([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8196)]
		void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16388)]
		void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32776)]
		void MasterAdded([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8200)]
		void MasterChanged([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16392)]
		void BeforeMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("page", typeof(VisioApi.IVPage))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32784)]
		void PageAdded([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("page", typeof(VisioApi.IVPage))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8208)]
		void PageChanged([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("page", typeof(VisioApi.IVPage))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16400)]
		void BeforePageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page);

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
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void RunModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void DesignModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void BeforeDocumentSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void BeforeDocumentSaveAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

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
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void QueryCancelDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void DocumentCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(300)]
		void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(301)]
		void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(400)]
		void QueryCancelMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(401)]
		void MasterDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("page", typeof(VisioApi.IVPage))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(500)]
		void QueryCancelPageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("page", typeof(VisioApi.IVPage))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(501)]
		void PageDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object page);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(802)]
		void ShapeParentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.Shape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(803)]
		void BeforeShapeTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.Shape))]
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

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("dataRecordset", typeof(VisioApi.IVDataRecordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16416)]
		void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("dataRecordsetChanged", typeof(VisioApi.IVDataRecordsetChangedEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8224)]
		void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("dataRecordset", typeof(VisioApi.IVDataRecordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32800)]
		void DataRecordsetAdded([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [SinkArgument("dataRecordsetID", SinkArgumentType.Int32)]
        [SinkArgument("dataRowID", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(805)]
		void ShapeLinkAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("shape", typeof(VisioApi.IVShape))]
        [SinkArgument("dataRecordsetID", SinkArgumentType.Int32)]
        [SinkArgument("dataRowID", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(806)]
		void ShapeLinkDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void AfterRemoveHiddenInformation([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Visio", 14,15,16)]
        [SinkArgument("shapePair", typeof(VisioApi.IVRelatedShapePairEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(502)]
		void ContainerRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		[SupportByVersion("Visio", 14,15,16)]
        [SinkArgument("shapePair", typeof(VisioApi.IVRelatedShapePairEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(503)]
		void ContainerRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		[SupportByVersion("Visio", 14,15,16)]
        [SinkArgument("shapePair", typeof(VisioApi.IVRelatedShapePairEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(504)]
		void CalloutRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		[SupportByVersion("Visio", 14,15,16)]
        [SinkArgument("shapePair", typeof(VisioApi.IVRelatedShapePairEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(505)]
		void CalloutRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair);

		[SupportByVersion("Visio", 14,15,16)]
        [SinkArgument("ruleSet", typeof(VisioApi.IVValidationRuleSet))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void RuleSetValidated([In, MarshalAs(UnmanagedType.IDispatch)] object ruleSet);

		[SupportByVersion("Visio", 15, 16)]
        [SinkArgument("replaceShapes", typeof(VisioApi.IVReplaceShapesEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(911)]
		void QueryCancelReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes);

        [SupportByVersion("Visio", 15, 16)]
        [SinkArgument("replaceShapes", typeof(VisioApi.IVReplaceShapesEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(912)]
		void ReplaceShapesCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes);

		[SupportByVersion("Visio", 15, 16)]
        [SinkArgument("replaceShapes", typeof(VisioApi.IVReplaceShapesEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(913)]
		void BeforeReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes);

		[SupportByVersion("Visio", 15, 16)]
        [SinkArgument("sel", typeof(VisioApi.IVSelection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(914)]
		void AfterReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object sel);

		[SupportByVersion("Visio", 15, 16)]
        [SinkArgument("coauthMergeObjects", typeof(VisioApi.IVCoauthMergeEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void AfterDocumentMerge([In, MarshalAs(UnmanagedType.IDispatch)] object coauthMergeObjects);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EDocuments_SinkHelper : SinkHelper, EDocuments
	{
		#region Static
		
		public static readonly string Id = "000D0B03-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public EDocuments_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region EDocuments
		
        public void DocumentOpened([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("DocumentOpened"))
            {
				Invoker.ReleaseParamsArray(doc);
				return;
			}

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("DocumentOpened", ref paramsArray);
		}

        public void DocumentCreated([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("DocumentCreated"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("DocumentCreated", ref paramsArray);
		}

        public void DocumentSaved([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("DocumentSaved"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("DocumentSaved", ref paramsArray);
		}

        public void DocumentSavedAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("DocumentSavedAs"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("DocumentSavedAs", ref paramsArray);
		}

        public void DocumentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("DocumentChanged"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("DocumentChanged", ref paramsArray);
		}

        public void BeforeDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("BeforeDocumentClose"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("BeforeDocumentClose", ref paramsArray);
		}

        public void StyleAdded([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
            if (!Validate("StyleAdded"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("StyleAdded", ref paramsArray);
		}

        public void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
            if (!Validate("StyleChanged"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("StyleChanged", ref paramsArray);
		}

        public void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
            if (!Validate("BeforeStyleDelete"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("BeforeStyleDelete", ref paramsArray);
		}

        public void MasterAdded([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
            if (!Validate("MasterAdded"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			EventBinding.RaiseCustomEvent("MasterAdded", ref paramsArray);
		}

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

        public void PageAdded([In, MarshalAs(UnmanagedType.IDispatch)] object page)
		{
            if (!Validate("PageAdded"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
			paramsArray[0] = newPage;
			EventBinding.RaiseCustomEvent("PageAdded", ref paramsArray);
		}

        public void PageChanged([In, MarshalAs(UnmanagedType.IDispatch)] object page)
		{
            if (!Validate("PageChanged"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
			paramsArray[0] = newPage;
			EventBinding.RaiseCustomEvent("PageChanged", ref paramsArray);
		}

        public void BeforePageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page)
		{
            if (!Validate("BeforePageDelete"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
			paramsArray[0] = newPage;
			EventBinding.RaiseCustomEvent("BeforePageDelete", ref paramsArray);
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

        public void RunModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("CellChanged"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("RunModeEntered", ref paramsArray);
		}

        public void DesignModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("DesignModeEntered"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("DesignModeEntered", ref paramsArray);
		}

        public void BeforeDocumentSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("BeforeDocumentSave"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("BeforeDocumentSave", ref paramsArray);
		}

        public void BeforeDocumentSaveAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("BeforeDocumentSaveAs"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("BeforeDocumentSaveAs", ref paramsArray);
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

        public void QueryCancelDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("QueryCancelDocumentClose"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("QueryCancelDocumentClose", ref paramsArray);
		}

        public void DocumentCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("DocumentCloseCanceled"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("DocumentCloseCanceled", ref paramsArray);
		}

        public void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
            if (!Validate("QueryCancelStyleDelete"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("QueryCancelStyleDelete", ref paramsArray);
		}

        public void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
            if (!Validate("StyleDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("StyleDeleteCanceled", ref paramsArray);
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

        public void QueryCancelPageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page)
		{
            if (!Validate("QueryCancelPageDelete"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
			paramsArray[0] = newPage;
			EventBinding.RaiseCustomEvent("QueryCancelPageDelete", ref paramsArray);
		}

        public void PageDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object page)
		{
            if (!Validate("PageDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
			paramsArray[0] = newPage;
			EventBinding.RaiseCustomEvent("PageDeleteCanceled", ref paramsArray);
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
            if (!Validate("UngroupCanceled"))
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

        public void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset)
		{
            if (!Validate("BeforeDataRecordsetDelete"))
            {
                Invoker.ReleaseParamsArray(dataRecordset);
                return;
            }

            NetOffice.VisioApi.IVDataRecordset newDataRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, dataRecordset) as NetOffice.VisioApi.IVDataRecordset;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordset;
			EventBinding.RaiseCustomEvent("BeforeDataRecordsetDelete", ref paramsArray);
		}

        public void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged)
		{
            if (!Validate("DataRecordsetChanged"))
            {
                Invoker.ReleaseParamsArray(dataRecordsetChanged);
                return;
            }

            NetOffice.VisioApi.IVDataRecordsetChangedEvent newDataRecordsetChanged = Factory.CreateEventArgumentObjectFromComProxy(EventClass, dataRecordsetChanged) as NetOffice.VisioApi.IVDataRecordsetChangedEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordsetChanged;
			EventBinding.RaiseCustomEvent("DataRecordsetChanged", ref paramsArray);
		}

        public void DataRecordsetAdded([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset)
		{
            if (!Validate("DataRecordsetAdded"))
            {
                Invoker.ReleaseParamsArray(dataRecordset);
                return;
            }

            NetOffice.VisioApi.IVDataRecordset newDataRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, dataRecordset) as NetOffice.VisioApi.IVDataRecordset;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordset;
			EventBinding.RaiseCustomEvent("DataRecordsetAdded", ref paramsArray);
		}

        public void ShapeLinkAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID)
		{
            if (!Validate("ShapeLinkAdded"))
            {
                Invoker.ReleaseParamsArray(shape, dataRecordsetID, dataRowID);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            Int32 newDataRecordsetID = ToInt32(dataRecordsetID);
			Int32 newDataRowID = ToInt32(dataRowID);
			object[] paramsArray = new object[3];
			paramsArray[0] = newShape;
			paramsArray[1] = newDataRecordsetID;
			paramsArray[2] = newDataRowID;
			EventBinding.RaiseCustomEvent("ShapeLinkAdded", ref paramsArray);
		}

        public void ShapeLinkDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID)
		{
            if (!Validate("ShapeLinkDeleted"))
            {
                Invoker.ReleaseParamsArray(shape, dataRecordsetID, dataRowID);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            Int32 newDataRecordsetID = ToInt32(dataRecordsetID);
			Int32 newDataRowID = ToInt32(dataRowID);
			object[] paramsArray = new object[3];
			paramsArray[0] = newShape;
			paramsArray[1] = newDataRecordsetID;
			paramsArray[2] = newDataRowID;
			EventBinding.RaiseCustomEvent("ShapeLinkDeleted", ref paramsArray);
		}

        public void AfterRemoveHiddenInformation([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("AfterRemoveHiddenInformation"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument;
            object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			EventBinding.RaiseCustomEvent("AfterRemoveHiddenInformation", ref paramsArray);
		}

        public void ContainerRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
            if (!Validate("ContainerRelationshipAdded"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			EventBinding.RaiseCustomEvent("ContainerRelationshipAdded", ref paramsArray);
		}

        public void ContainerRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
            if (!Validate("ContainerRelationshipDeleted"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			EventBinding.RaiseCustomEvent("ContainerRelationshipDeleted", ref paramsArray);
        }

        public void CalloutRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
            if (!Validate("CalloutRelationshipAdded"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			EventBinding.RaiseCustomEvent("CalloutRelationshipAdded", ref paramsArray);
		}

        public void CalloutRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
		{
            if (!Validate("CalloutRelationshipDeleted"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newShapePair;
			EventBinding.RaiseCustomEvent("CalloutRelationshipDeleted", ref paramsArray);
		}

        public void RuleSetValidated([In, MarshalAs(UnmanagedType.IDispatch)] object ruleSet)
		{
            if (!Validate("RuleSetValidated"))
            {
                Invoker.ReleaseParamsArray(ruleSet);
                return;
            }

            NetOffice.VisioApi.IVValidationRuleSet newRuleSet = Factory.CreateEventArgumentObjectFromComProxy(EventClass, ruleSet) as NetOffice.VisioApi.IVValidationRuleSet;
            object[] paramsArray = new object[1];
			paramsArray[0] = newRuleSet;
			EventBinding.RaiseCustomEvent("RuleSetValidated", ref paramsArray);
		}

        public void QueryCancelReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
		{
            if (!Validate("QueryCancelReplaceShapes"))
            {
                Invoker.ReleaseParamsArray(replaceShapes);
                return;
            }

            NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateEventArgumentObjectFromComProxy(EventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newreplaceShapes;
			EventBinding.RaiseCustomEvent("QueryCancelReplaceShapes", ref paramsArray);
		}

        public void ReplaceShapesCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
		{
            if (!Validate("ReplaceShapesCanceled"))
            {
                Invoker.ReleaseParamsArray(replaceShapes);
                return;
            }

            NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateEventArgumentObjectFromComProxy(EventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newreplaceShapes;
			EventBinding.RaiseCustomEvent("ReplaceShapesCanceled", ref paramsArray);
		}

        public void BeforeReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
		{
            if (!Validate("BeforeReplaceShapes"))
            {
                Invoker.ReleaseParamsArray(replaceShapes);
                return;
            }

            NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateEventArgumentObjectFromComProxy(EventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newreplaceShapes;
			EventBinding.RaiseCustomEvent("BeforeReplaceShapes", ref paramsArray);
		}
        
        public void AfterReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object sel)
		{
            if (!Validate("AfterReplaceShapes"))
            {
                Invoker.ReleaseParamsArray(sel);
                return;
            }

            NetOffice.VisioApi.IVSelection newsel = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sel) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
			paramsArray[0] = newsel;
			EventBinding.RaiseCustomEvent("AfterReplaceShapes", ref paramsArray);
		}
        
        public void AfterDocumentMerge([In, MarshalAs(UnmanagedType.IDispatch)] object coauthMergeObjects)
		{
            if (!Validate("AfterDocumentMerge"))
            {
                Invoker.ReleaseParamsArray(coauthMergeObjects);
                return;
            }

            NetOffice.VisioApi.IVCoauthMergeEvent newcoauthMergeObjects = Factory.CreateEventArgumentObjectFromComProxy(EventClass, coauthMergeObjects) as NetOffice.VisioApi.IVCoauthMergeEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newcoauthMergeObjects;
			EventBinding.RaiseCustomEvent("AfterDocumentMerge", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}