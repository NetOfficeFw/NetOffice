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
    [ComImport, Guid("000D0B00-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EApplication
	{
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4097)]
		void AppActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4098)]
		void AppDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4100)]
		void AppObjActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4104)]
		void AppObjDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4112)]
		void BeforeQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4128)]
		void BeforeModal([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4160)]
		void AfterModal([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32769)]
		void WindowOpened([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(701)]
		void SelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16385)]
		void BeforeWindowClosed([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4224)]
		void WindowActivated([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(702)]
		void BeforeWindowSelDelete([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(703)]
		void BeforeWindowPageTurn([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(704)]
		void WindowTurnedToPage([In, MarshalAs(UnmanagedType.IDispatch)] object window);

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
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [SinkArgument("sequenceNum", SinkArgumentType.Int32)]
        [SinkArgument("contextString", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4352)]
		void MarkerEvent([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object sequenceNum, [In] object contextString);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4608)]
		void NoEventsPending([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5120)]
		void VisioIsIdle([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(200)]
		void MustFlushScopeBeginning([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(201)]
		void MustFlushScopeEnded([In, MarshalAs(UnmanagedType.IDispatch)] object app);

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
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [SinkArgument("nScopeID", SinkArgumentType.Int32)]
        [SinkArgument("bstrDescription", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(202)]
		void EnterScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [SinkArgument("nScopeID", SinkArgumentType.Int32)]
        [SinkArgument("bstrDescription", SinkArgumentType.String)]
        [SinkArgument("bErrOrCancelled", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(203)]
		void ExitScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription, [In] object bErrOrCancelled);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(204)]
		void QueryCancelQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(205)]
		void QuitCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8193)]
		void WindowChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(705)]
		void ViewChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(706)]
		void QueryCancelWindowClose([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(707)]
		void WindowCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object window);

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

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(206)]
		void QueryCancelSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(207)]
		void SuspendCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(208)]
		void BeforeSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(209)]
		void AfterResume([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("mSG", typeof(VisioApi.IVMSGWrap))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(708)]
		void OnKeystrokeMessageForAddon([In, MarshalAs(UnmanagedType.IDispatch)] object mSG);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(709)]
		void MouseDown([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(710)]
		void MouseMove([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(711)]
		void MouseUp([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(712)]
		void KeyDown([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyAscii", SinkArgumentType.Int32)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(713)]
		void KeyPress([In] object keyAscii, [In] [Out] ref object cancelDefault);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(714)]
		void KeyUp([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(210)]
		void QueryCancelSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(211)]
		void SuspendEventsCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(212)]
		void BeforeSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(213)]
		void AfterResumeEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app);

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
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class EApplication_SinkHelper : SinkHelper, EApplication
    {
        #region Static

        public static readonly string Id = "000D0B00-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public EApplication_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region EApplication Members

        public void AppActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AppActivated"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AppActivated", ref paramsArray);
        }

        public void AppDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AppDeactivated"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AppDeactivated", ref paramsArray);
        }

        public void AppObjActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AppObjActivated"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AppObjActivated", ref paramsArray);
        }

        public void AppObjDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AppObjDeactivated"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AppObjDeactivated", ref paramsArray);
        }

        public void BeforeQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("BeforeQuit"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("BeforeQuit", ref paramsArray);
        }

        public void BeforeModal([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("BeforeModal"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("BeforeModal", ref paramsArray);
        }

        public void AfterModal([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AfterModal"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AfterModal", ref paramsArray);
        }

        public void WindowOpened([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowOpened"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowOpened", ref paramsArray);
        }

        public void SelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("SelectionChanged"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("SelectionChanged", ref paramsArray);
        }

        public void BeforeWindowClosed([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("BeforeWindowClosed"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("BeforeWindowClosed", ref paramsArray);
        }

        public void WindowActivated([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowActivated"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowActivated", ref paramsArray);
        }

        public void BeforeWindowSelDelete([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("BeforeWindowSelDelete"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("BeforeWindowSelDelete", ref paramsArray);
        }

        public void BeforeWindowPageTurn([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("BeforeWindowPageTurn"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("BeforeWindowPageTurn", ref paramsArray);
        }

        public void WindowTurnedToPage([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowTurnedToPage"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowTurnedToPage", ref paramsArray);
        }

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

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument; object[] paramsArray = new object[1];
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

            NetOffice.VisioApi.IVDocument newdoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.VisioApi.IVDocument; object[] paramsArray = new object[1];
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

        public void MarkerEvent([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object sequenceNum, [In] object contextString)
        {
            if (!Validate("MarkerEvent"))
            {
                Invoker.ReleaseParamsArray(app, sequenceNum, contextString);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            Int32 newSequenceNum = ToInt32(sequenceNum);
            string newContextString = ToString(contextString);
            object[] paramsArray = new object[3];
            paramsArray[0] = newapp;
            paramsArray[1] = newSequenceNum;
            paramsArray[2] = newContextString;
            EventBinding.RaiseCustomEvent("MarkerEvent", ref paramsArray);
        }

        public void NoEventsPending([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("NoEventsPending"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("NoEventsPending", ref paramsArray);
        }

        public void VisioIsIdle([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("VisioIsIdle"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("VisioIsIdle", ref paramsArray);
        }

        public void MustFlushScopeBeginning([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("MustFlushScopeBeginning"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("MustFlushScopeBeginning", ref paramsArray);
        }

        public void MustFlushScopeEnded([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("MustFlushScopeEnded"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("MustFlushScopeEnded", ref paramsArray);
        }

        public void RunModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("RunModeEntered"))
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

        public void EnterScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription)
        {
            if (!Validate("EnterScope"))
            {
                Invoker.ReleaseParamsArray(app, nScopeID, bstrDescription);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            Int32 newnScopeID = ToInt32(nScopeID);
            string newbstrDescription = ToString(bstrDescription);
            object[] paramsArray = new object[3];
            paramsArray[0] = newapp;
            paramsArray[1] = newnScopeID;
            paramsArray[2] = newbstrDescription;
            EventBinding.RaiseCustomEvent("EnterScope", ref paramsArray);
        }

        public void ExitScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription, [In] object bErrOrCancelled)
        {
            if (!Validate("ExitScope"))
            {
                Invoker.ReleaseParamsArray(app, nScopeID, bstrDescription, bErrOrCancelled);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            Int32 newnScopeID = ToInt32(nScopeID);
            string newbstrDescription = ToString(bstrDescription);
            bool newbErrOrCancelled = ToBoolean(bErrOrCancelled);
            object[] paramsArray = new object[4];
            paramsArray[0] = newapp;
            paramsArray[1] = newnScopeID;
            paramsArray[2] = newbstrDescription;
            paramsArray[3] = newbErrOrCancelled;
            EventBinding.RaiseCustomEvent("ExitScope", ref paramsArray);
        }

        public void QueryCancelQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("QueryCancelQuit"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("QueryCancelQuit", ref paramsArray);
        }

        public void QuitCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("QuitCanceled"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("QuitCanceled", ref paramsArray);
        }

        public void WindowChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowChanged"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowChanged", ref paramsArray);
        }

        public void ViewChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("ViewChanged"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("ViewChanged", ref paramsArray);
        }

        public void QueryCancelWindowClose([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("QueryCancelWindowClose"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("QueryCancelWindowClose", ref paramsArray);
        }

        public void WindowCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object window)
        {
            if (!Validate("WindowCloseCanceled"))
            {
                Invoker.ReleaseParamsArray(window);
                return;
            }

            NetOffice.VisioApi.IVWindow newWindow = Factory.CreateEventArgumentObjectFromComProxy(EventClass, window) as NetOffice.VisioApi.IVWindow;
            object[] paramsArray = new object[1];
            paramsArray[0] = newWindow;
            EventBinding.RaiseCustomEvent("WindowCloseCanceled", ref paramsArray);
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

        public void QueryCancelSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("QueryCancelSuspend"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("QueryCancelSuspend", ref paramsArray);
        }

        public void SuspendCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("SuspendCanceled"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("SuspendCanceled", ref paramsArray);
        }

        public void BeforeSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("BeforeSuspend"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("BeforeSuspend", ref paramsArray);
        }

        public void AfterResume([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("AfterResume"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
            paramsArray[0] = newapp;
            EventBinding.RaiseCustomEvent("AfterResume", ref paramsArray);
        }

        public void OnKeystrokeMessageForAddon([In, MarshalAs(UnmanagedType.IDispatch)] object mSG)
        {
            if (!Validate("OnKeystrokeMessageForAddon"))
            {
                Invoker.ReleaseParamsArray(mSG);
                return;
            }

            NetOffice.VisioApi.IVMSGWrap newMSG = Factory.CreateEventArgumentObjectFromComProxy(EventClass, mSG) as NetOffice.VisioApi.IVMSGWrap;
            object[] paramsArray = new object[1];
            paramsArray[0] = newMSG;
            EventBinding.RaiseCustomEvent("OnKeystrokeMessageForAddon", ref paramsArray);
        }

        public void MouseDown([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
                return;
            }

			Int32 newButton = ToInt32(button);
			Int32 newKeyButtonState = ToInt32(keyButtonState);
			Double newx = ToDouble(x);
			Double newy = ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[4]);
		}

        public void MouseMove([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
                return;
            }

            Int32 newButton = ToInt32(button);
			Int32 newKeyButtonState = ToInt32(keyButtonState);
			Double newx = ToDouble(x);
			Double newy = ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[4]);
		}

        public void MouseUp([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
                return;
            }

            Int32 newButton = ToInt32(button);
			Int32 newKeyButtonState = ToInt32(keyButtonState);
			Double newx = ToDouble(x);
			Double newy = ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[4]);
		}

        public void KeyDown([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("KeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, keyButtonState, cancelDefault);
                return;
            }

			Int32 newKeyCode = ToInt32(keyCode);
			Int32 newKeyButtonState = ToInt32(keyButtonState);
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newKeyButtonState;
			paramsArray.SetValue(cancelDefault, 2);
			EventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[2]);
		}

        public void KeyPress([In] object keyAscii, [In] [Out] ref object cancelDefault)
		{
            if (!Validate("KeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii, cancelDefault);
                return;
            }

			Int32 newKeyAscii = ToInt32(keyAscii);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyAscii;
			paramsArray.SetValue(cancelDefault, 1);
			EventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[1]);
        }

        public void KeyUp([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("KeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, keyButtonState, cancelDefault);
                return;
            }

			Int32 newKeyCode = ToInt32(keyCode);
			Int32 newKeyButtonState = ToInt32(keyButtonState);
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newKeyButtonState;
			paramsArray.SetValue(cancelDefault, 2);
			EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);

			cancelDefault = ToBoolean(paramsArray[2]);
        }

        public void QueryCancelSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("QueryCancelSuspendEvents"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			EventBinding.RaiseCustomEvent("QueryCancelSuspendEvents", ref paramsArray);
		}

        public void SuspendEventsCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
            if (!Validate("SuspendEventsCanceled"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			EventBinding.RaiseCustomEvent("SuspendEventsCanceled", ref paramsArray);
		}

        public void BeforeSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app)
        {
            if (!Validate("BeforeSuspendEvents"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			EventBinding.RaiseCustomEvent("BeforeSuspendEvents", ref paramsArray);
		}

        public void AfterResumeEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
            if (!Validate("AfterResumeEvents"))
            {
                Invoker.ReleaseParamsArray(app);
                return;
            }

            NetOffice.VisioApi.IVApplication newapp = Factory.CreateEventArgumentObjectFromComProxy(EventClass, app) as NetOffice.VisioApi.IVApplication;
            object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			EventBinding.RaiseCustomEvent("AfterResumeEvents", ref paramsArray);
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
            if (!Validate("BeforeDataRecordsetDelete"))
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

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}