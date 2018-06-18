using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.EventContracts
{
    /// <summary>
    /// EApplication
    /// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000D0B00-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EApplication
	{
		/// <summary>
		/// AppActivated
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4097)]
		void AppActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// AppDeactivated
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4098)]
		void AppDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// AppObjActivated
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4100)]
		void AppObjActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// AppObjDeactivated
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4104)]
		void AppObjDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// BeforeQuit
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4112)]
		void BeforeQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// BeforeModal
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4128)]
		void BeforeModal([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// AfterModal
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4160)]
		void AfterModal([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// WindowOpened
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32769)]
		void WindowOpened([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// SelectionChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(701)]
		void SelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// BeforeWindowClosed
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16385)]
		void BeforeWindowClosed([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// WindowActivated
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4224)]
		void WindowActivated([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// BeforeWindowSelDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(702)]
		void BeforeWindowSelDelete([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// BeforeWindowPageTurn
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(703)]
		void BeforeWindowPageTurn([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// WindowTurnedToPage
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(704)]
		void WindowTurnedToPage([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// DocumentOpened
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void DocumentOpened([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// DocumentCreated
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void DocumentCreated([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// DocumentSaved
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void DocumentSaved([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// DocumentSavedAs
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void DocumentSavedAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// DocumentChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8194)]
		void DocumentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// BeforeDocumentClose
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16386)]
		void BeforeDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// StyleAdded
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32772)]
		void StyleAdded([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		/// <summary>
		/// StyleChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8196)]
		void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		/// <summary>
		/// BeforeStyleDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16388)]
		void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		/// <summary>
		/// MasterAdded
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32776)]
		void MasterAdded([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		/// <summary>
		/// MasterChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8200)]
		void MasterChanged([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		/// <summary>
		/// BeforeMasterDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16392)]
		void BeforeMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master);

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
		/// MarkerEvent
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [SinkArgument("sequenceNum", SinkArgumentType.Int32)]
        [SinkArgument("contextString", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4352)]
		void MarkerEvent([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object sequenceNum, [In] object contextString);

		/// <summary>
		/// NoEventsPending
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4608)]
		void NoEventsPending([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// VisioIsIdle
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5120)]
		void VisioIsIdle([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// MustFlushScopeBeginning
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(200)]
		void MustFlushScopeBeginning([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// MustFlushScopeEnded
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(201)]
		void MustFlushScopeEnded([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// RunModeEntered
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void RunModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// DesignModeEntered
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void DesignModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// BeforeDocumentSave
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void BeforeDocumentSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// BeforeDocumentSaveAs
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void BeforeDocumentSaveAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

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
		/// EnterScope
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [SinkArgument("nScopeID", SinkArgumentType.Int32)]
        [SinkArgument("bstrDescription", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(202)]
		void EnterScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription);

		/// <summary>
		/// ExitScope
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [SinkArgument("nScopeID", SinkArgumentType.Int32)]
        [SinkArgument("bstrDescription", SinkArgumentType.String)]
        [SinkArgument("bErrOrCancelled", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(203)]
		void ExitScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription, [In] object bErrOrCancelled);

		/// <summary>
		/// QueryCancelQuit
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(204)]
		void QueryCancelQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// QuitCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(205)]
		void QuitCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// WindowChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8193)]
		void WindowChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// ViewChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(705)]
		void ViewChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// QueryCancelWindowClose
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(706)]
		void QueryCancelWindowClose([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// WindowCloseCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("window", typeof(VisioApi.IVWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(707)]
		void WindowCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		/// <summary>
		/// QueryCancelDocumentClose
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void QueryCancelDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// DocumentCloseCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void DocumentCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// QueryCancelStyleDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(300)]
		void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		/// <summary>
		/// StyleDeleteCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(301)]
		void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		/// <summary>
		/// QueryCancelMasterDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(400)]
		void QueryCancelMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		/// <summary>
		/// MasterDeleteCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("master", typeof(VisioApi.IVMaster))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(401)]
		void MasterDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object master);

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
		/// QueryCancelSuspend
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(206)]
		void QueryCancelSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// SuspendCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(207)]
		void SuspendCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// BeforeSuspend
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(208)]
		void BeforeSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// AfterResume
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(209)]
		void AfterResume([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// OnKeystrokeMessageForAddon
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("mSG", typeof(VisioApi.IVMSGWrap))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(708)]
		void OnKeystrokeMessageForAddon([In, MarshalAs(UnmanagedType.IDispatch)] object mSG);

		/// <summary>
		/// MouseDown
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(709)]
		void MouseDown([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		/// <summary>
		/// MouseMove
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(710)]
		void MouseMove([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		/// <summary>
		/// MouseUp
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Double)]
        [SinkArgument("y", SinkArgumentType.Double)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(711)]
		void MouseUp([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		/// <summary>
		/// KeyDown
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(712)]
		void KeyDown([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);

		/// <summary>
		/// KeyPress
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyAscii", SinkArgumentType.Int32)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(713)]
		void KeyPress([In] object keyAscii, [In] [Out] ref object cancelDefault);

		/// <summary>
		/// KeyUp
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("keyCode", SinkArgumentType.Int32)]
        [SinkArgument("keyButtonState", SinkArgumentType.Int32)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(714)]
		void KeyUp([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);

		/// <summary>
		/// QueryCancelSuspendEvents
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(210)]
		void QueryCancelSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// SuspendEventsCanceled
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(211)]
		void SuspendEventsCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// BeforeSuspendEvents
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(212)]
		void BeforeSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		/// <summary>
		/// AfterResumeEvents
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("app", typeof(VisioApi.IVApplication))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(213)]
		void AfterResumeEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app);

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
		/// BeforeDataRecordsetDelete
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("dataRecordset", typeof(VisioApi.IVDataRecordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16416)]
		void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset);

		/// <summary>
		/// DataRecordsetChanged
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("dataRecordsetChanged", typeof(VisioApi.IVDataRecordsetChangedEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8224)]
		void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged);

		/// <summary>
		/// DataRecordsetAdded
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("dataRecordset", typeof(VisioApi.IVDataRecordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32800)]
		void DataRecordsetAdded([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset);

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
		/// AfterRemoveHiddenInformation
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("doc", typeof(VisioApi.IVDocument))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void AfterRemoveHiddenInformation([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

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
		/// RuleSetValidated
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
        [SinkArgument("ruleSet", typeof(VisioApi.IVValidationRuleSet))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void RuleSetValidated([In, MarshalAs(UnmanagedType.IDispatch)] object ruleSet);

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
