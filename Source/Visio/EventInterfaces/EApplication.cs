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
	[ComImport, Guid("000D0B00-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EApplication
	{
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4097)]
		void AppActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4098)]
		void AppDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4100)]
		void AppObjActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4104)]
		void AppObjDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4112)]
		void BeforeQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4128)]
		void BeforeModal([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4160)]
		void AfterModal([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32769)]
		void WindowOpened([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(701)]
		void SelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16385)]
		void BeforeWindowClosed([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4224)]
		void WindowActivated([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(702)]
		void BeforeWindowSelDelete([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(703)]
		void BeforeWindowPageTurn([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(704)]
		void WindowTurnedToPage([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void DocumentOpened([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void DocumentCreated([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void DocumentSaved([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void DocumentSavedAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8194)]
		void DocumentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16386)]
		void BeforeDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32772)]
		void StyleAdded([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8196)]
		void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16388)]
		void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32776)]
		void MasterAdded([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8200)]
		void MasterChanged([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16392)]
		void BeforeMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32784)]
		void PageAdded([In, MarshalAs(UnmanagedType.IDispatch)] object page);

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
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4352)]
		void MarkerEvent([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object sequenceNum, [In] object contextString);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4608)]
		void NoEventsPending([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5120)]
		void VisioIsIdle([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(200)]
		void MustFlushScopeBeginning([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(201)]
		void MustFlushScopeEnded([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void RunModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void DesignModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void BeforeDocumentSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void BeforeDocumentSaveAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

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
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(202)]
		void EnterScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(203)]
		void ExitScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription, [In] object bErrOrCancelled);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(204)]
		void QueryCancelQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(205)]
		void QuitCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8193)]
		void WindowChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(705)]
		void ViewChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(706)]
		void QueryCancelWindowClose([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(707)]
		void WindowCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object window);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void QueryCancelDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void DocumentCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(300)]
		void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(301)]
		void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(400)]
		void QueryCancelMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(401)]
		void MasterDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object master);

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

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(206)]
		void QueryCancelSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(207)]
		void SuspendCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(208)]
		void BeforeSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(209)]
		void AfterResume([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(708)]
		void OnKeystrokeMessageForAddon([In, MarshalAs(UnmanagedType.IDispatch)] object mSG);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(709)]
		void MouseDown([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(710)]
		void MouseMove([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(711)]
		void MouseUp([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(712)]
		void KeyDown([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(713)]
		void KeyPress([In] object keyAscii, [In] [Out] ref object cancelDefault);

		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(714)]
		void KeyUp([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(210)]
		void QueryCancelSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(211)]
		void SuspendEventsCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(212)]
		void BeforeSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(213)]
		void AfterResumeEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app);

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
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16416)]
		void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8224)]
		void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32800)]
		void DataRecordsetAdded([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(805)]
		void ShapeLinkAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(806)]
		void ShapeLinkDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void AfterRemoveHiddenInformation([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

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

		[SupportByVersionAttribute("Visio", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void RuleSetValidated([In, MarshalAs(UnmanagedType.IDispatch)] object ruleSet);

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
	public class EApplication_SinkHelper : SinkHelper, EApplication
	{
		#region Static
		
		public static readonly string Id = "000D0B00-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public EApplication_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region EApplication Members
		
		public void AppActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AppActivated");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("AppActivated", ref paramsArray);
		}

		public void AppDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AppDeactivated");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("AppDeactivated", ref paramsArray);
		}

		public void AppObjActivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AppObjActivated");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("AppObjActivated", ref paramsArray);
		}

		public void AppObjDeactivated([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AppObjDeactivated");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("AppObjDeactivated", ref paramsArray);
		}

		public void BeforeQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeQuit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("BeforeQuit", ref paramsArray);
		}

		public void BeforeModal([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeModal");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("BeforeModal", ref paramsArray);
		}

		public void AfterModal([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterModal");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("AfterModal", ref paramsArray);
		}

		public void WindowOpened([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowOpened");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("WindowOpened", ref paramsArray);
		}

		public void SelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SelectionChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("SelectionChanged", ref paramsArray);
		}

		public void BeforeWindowClosed([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeWindowClosed");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("BeforeWindowClosed", ref paramsArray);
		}

		public void WindowActivated([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowActivated");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("WindowActivated", ref paramsArray);
		}

		public void BeforeWindowSelDelete([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeWindowSelDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("BeforeWindowSelDelete", ref paramsArray);
		}

		public void BeforeWindowPageTurn([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeWindowPageTurn");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("BeforeWindowPageTurn", ref paramsArray);
		}

		public void WindowTurnedToPage([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowTurnedToPage");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("WindowTurnedToPage", ref paramsArray);
		}

		public void DocumentOpened([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentOpened");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("DocumentOpened", ref paramsArray);
		}

		public void DocumentCreated([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentCreated");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("DocumentCreated", ref paramsArray);
		}

		public void DocumentSaved([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentSaved");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("DocumentSaved", ref paramsArray);
		}

		public void DocumentSavedAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentSavedAs");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("DocumentSavedAs", ref paramsArray);
		}

		public void DocumentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("DocumentChanged", ref paramsArray);
		}

		public void BeforeDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDocumentClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("BeforeDocumentClose", ref paramsArray);
		}

		public void StyleAdded([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("StyleAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(style);
				return;
			}

			NetOffice.VisioApi.IVStyle newStyle = Factory.CreateObjectFromComProxy(_eventClass, style) as NetOffice.VisioApi.IVStyle;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			_eventBinding.RaiseCustomEvent("StyleAdded", ref paramsArray);
		}

		public void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("StyleChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(style);
				return;
			}

			NetOffice.VisioApi.IVStyle newStyle = Factory.CreateObjectFromComProxy(_eventClass, style) as NetOffice.VisioApi.IVStyle;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			_eventBinding.RaiseCustomEvent("StyleChanged", ref paramsArray);
		}

		public void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeStyleDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(style);
				return;
			}

			NetOffice.VisioApi.IVStyle newStyle = Factory.CreateObjectFromComProxy(_eventClass, style) as NetOffice.VisioApi.IVStyle;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			_eventBinding.RaiseCustomEvent("BeforeStyleDelete", ref paramsArray);
		}

		public void MasterAdded([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MasterAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(master);
				return;
			}

			NetOffice.VisioApi.IVMaster newMaster = Factory.CreateObjectFromComProxy(_eventClass, master) as NetOffice.VisioApi.IVMaster;
			object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			_eventBinding.RaiseCustomEvent("MasterAdded", ref paramsArray);
		}

		public void MasterChanged([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MasterChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(master);
				return;
			}

			NetOffice.VisioApi.IVMaster newMaster = Factory.CreateObjectFromComProxy(_eventClass, master) as NetOffice.VisioApi.IVMaster;
			object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			_eventBinding.RaiseCustomEvent("MasterChanged", ref paramsArray);
		}

		public void BeforeMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeMasterDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(master);
				return;
			}

			NetOffice.VisioApi.IVMaster newMaster = Factory.CreateObjectFromComProxy(_eventClass, master) as NetOffice.VisioApi.IVMaster;
			object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			_eventBinding.RaiseCustomEvent("BeforeMasterDelete", ref paramsArray);
		}

		public void PageAdded([In, MarshalAs(UnmanagedType.IDispatch)] object page)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PageAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(page);
				return;
			}

			NetOffice.VisioApi.IVPage newPage = Factory.CreateObjectFromComProxy(_eventClass, page) as NetOffice.VisioApi.IVPage;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPage;
			_eventBinding.RaiseCustomEvent("PageAdded", ref paramsArray);
		}

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

		public void MarkerEvent([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object sequenceNum, [In] object contextString)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MarkerEvent");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app, sequenceNum, contextString);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			Int32 newSequenceNum = Convert.ToInt32(sequenceNum);
			string newContextString = Convert.ToString(contextString);
			object[] paramsArray = new object[3];
			paramsArray[0] = newapp;
			paramsArray[1] = newSequenceNum;
			paramsArray[2] = newContextString;
			_eventBinding.RaiseCustomEvent("MarkerEvent", ref paramsArray);
		}

		public void NoEventsPending([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NoEventsPending");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("NoEventsPending", ref paramsArray);
		}

		public void VisioIsIdle([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("VisioIsIdle");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("VisioIsIdle", ref paramsArray);
		}

		public void MustFlushScopeBeginning([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MustFlushScopeBeginning");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("MustFlushScopeBeginning", ref paramsArray);
		}

		public void MustFlushScopeEnded([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MustFlushScopeEnded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("MustFlushScopeEnded", ref paramsArray);
		}

		public void RunModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("RunModeEntered");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("RunModeEntered", ref paramsArray);
		}

		public void DesignModeEntered([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DesignModeEntered");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("DesignModeEntered", ref paramsArray);
		}

		public void BeforeDocumentSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDocumentSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("BeforeDocumentSave", ref paramsArray);
		}

		public void BeforeDocumentSaveAs([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDocumentSaveAs");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("BeforeDocumentSaveAs", ref paramsArray);
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

		public void EnterScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("EnterScope");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app, nScopeID, bstrDescription);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			Int32 newnScopeID = Convert.ToInt32(nScopeID);
			string newbstrDescription = Convert.ToString(bstrDescription);
			object[] paramsArray = new object[3];
			paramsArray[0] = newapp;
			paramsArray[1] = newnScopeID;
			paramsArray[2] = newbstrDescription;
			_eventBinding.RaiseCustomEvent("EnterScope", ref paramsArray);
		}

		public void ExitScope([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] object nScopeID, [In] object bstrDescription, [In] object bErrOrCancelled)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ExitScope");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app, nScopeID, bstrDescription, bErrOrCancelled);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			Int32 newnScopeID = Convert.ToInt32(nScopeID);
			string newbstrDescription = Convert.ToString(bstrDescription);
			bool newbErrOrCancelled = Convert.ToBoolean(bErrOrCancelled);
			object[] paramsArray = new object[4];
			paramsArray[0] = newapp;
			paramsArray[1] = newnScopeID;
			paramsArray[2] = newbstrDescription;
			paramsArray[3] = newbErrOrCancelled;
			_eventBinding.RaiseCustomEvent("ExitScope", ref paramsArray);
		}

		public void QueryCancelQuit([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelQuit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("QueryCancelQuit", ref paramsArray);
		}

		public void QuitCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QuitCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("QuitCanceled", ref paramsArray);
		}

		public void WindowChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("WindowChanged", ref paramsArray);
		}

		public void ViewChanged([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ViewChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("ViewChanged", ref paramsArray);
		}

		public void QueryCancelWindowClose([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelWindowClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("QueryCancelWindowClose", ref paramsArray);
		}

		public void WindowCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object window)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowCloseCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window);
				return;
			}

			NetOffice.VisioApi.IVWindow newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.VisioApi.IVWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWindow;
			_eventBinding.RaiseCustomEvent("WindowCloseCanceled", ref paramsArray);
		}

		public void QueryCancelDocumentClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelDocumentClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("QueryCancelDocumentClose", ref paramsArray);
		}

		public void DocumentCloseCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentCloseCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("DocumentCloseCanceled", ref paramsArray);
		}

		public void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelStyleDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(style);
				return;
			}

			NetOffice.VisioApi.IVStyle newStyle = Factory.CreateObjectFromComProxy(_eventClass, style) as NetOffice.VisioApi.IVStyle;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			_eventBinding.RaiseCustomEvent("QueryCancelStyleDelete", ref paramsArray);
		}

		public void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("StyleDeleteCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(style);
				return;
			}

			NetOffice.VisioApi.IVStyle newStyle = Factory.CreateObjectFromComProxy(_eventClass, style) as NetOffice.VisioApi.IVStyle;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			_eventBinding.RaiseCustomEvent("StyleDeleteCanceled", ref paramsArray);
		}

		public void QueryCancelMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelMasterDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(master);
				return;
			}

			NetOffice.VisioApi.IVMaster newMaster = Factory.CreateObjectFromComProxy(_eventClass, master) as NetOffice.VisioApi.IVMaster;
			object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			_eventBinding.RaiseCustomEvent("QueryCancelMasterDelete", ref paramsArray);
		}

		public void MasterDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MasterDeleteCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(master);
				return;
			}

			NetOffice.VisioApi.IVMaster newMaster = Factory.CreateObjectFromComProxy(_eventClass, master) as NetOffice.VisioApi.IVMaster;
			object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			_eventBinding.RaiseCustomEvent("MasterDeleteCanceled", ref paramsArray);
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

		public void QueryCancelSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelSuspend");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("QueryCancelSuspend", ref paramsArray);
		}

		public void SuspendCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SuspendCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("SuspendCanceled", ref paramsArray);
		}

		public void BeforeSuspend([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeSuspend");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("BeforeSuspend", ref paramsArray);
		}

		public void AfterResume([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterResume");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("AfterResume", ref paramsArray);
		}

		public void OnKeystrokeMessageForAddon([In, MarshalAs(UnmanagedType.IDispatch)] object mSG)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("OnKeystrokeMessageForAddon");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(mSG);
				return;
			}

			NetOffice.VisioApi.IVMSGWrap newMSG = Factory.CreateObjectFromComProxy(_eventClass, mSG) as NetOffice.VisioApi.IVMSGWrap;
			object[] paramsArray = new object[1];
			paramsArray[0] = newMSG;
			_eventBinding.RaiseCustomEvent("OnKeystrokeMessageForAddon", ref paramsArray);
		}

		public void MouseDown([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
				return;
			}

			Int32 newButton = Convert.ToInt32(button);
			Int32 newKeyButtonState = Convert.ToInt32(keyButtonState);
			Double newx = Convert.ToDouble(x);
			Double newy = Convert.ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			_eventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);

			cancelDefault = (bool)paramsArray[4];
		}

		public void MouseMove([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
				return;
			}

			Int32 newButton = Convert.ToInt32(button);
			Int32 newKeyButtonState = Convert.ToInt32(keyButtonState);
			Double newx = Convert.ToDouble(x);
			Double newy = Convert.ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			_eventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);

			cancelDefault = (bool)paramsArray[4];
		}

		public void MouseUp([In] object button, [In] object keyButtonState, [In] object x, [In] object y, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, keyButtonState, x, y, cancelDefault);
				return;
			}

			Int32 newButton = Convert.ToInt32(button);
			Int32 newKeyButtonState = Convert.ToInt32(keyButtonState);
			Double newx = Convert.ToDouble(x);
			Double newy = Convert.ToDouble(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newButton;
			paramsArray[1] = newKeyButtonState;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			paramsArray.SetValue(cancelDefault, 4);
			_eventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);

			cancelDefault = (bool)paramsArray[4];
		}

		public void KeyDown([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, keyButtonState, cancelDefault);
				return;
			}

			Int32 newKeyCode = Convert.ToInt32(keyCode);
			Int32 newKeyButtonState = Convert.ToInt32(keyButtonState);
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newKeyButtonState;
			paramsArray.SetValue(cancelDefault, 2);
			_eventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);

			cancelDefault = (bool)paramsArray[2];
		}

		public void KeyPress([In] object keyAscii, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyPress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyAscii, cancelDefault);
				return;
			}

			Int32 newKeyAscii = Convert.ToInt32(keyAscii);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyAscii;
			paramsArray.SetValue(cancelDefault, 1);
			_eventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);

			cancelDefault = (bool)paramsArray[1];
		}

		public void KeyUp([In] object keyCode, [In] object keyButtonState, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, keyButtonState, cancelDefault);
				return;
			}

			Int32 newKeyCode = Convert.ToInt32(keyCode);
			Int32 newKeyButtonState = Convert.ToInt32(keyButtonState);
			object[] paramsArray = new object[3];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newKeyButtonState;
			paramsArray.SetValue(cancelDefault, 2);
			_eventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);

			cancelDefault = (bool)paramsArray[2];
		}

		public void QueryCancelSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("QueryCancelSuspendEvents");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("QueryCancelSuspendEvents", ref paramsArray);
		}

		public void SuspendEventsCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SuspendEventsCanceled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("SuspendEventsCanceled", ref paramsArray);
		}

		public void BeforeSuspendEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeSuspendEvents");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("BeforeSuspendEvents", ref paramsArray);
		}

		public void AfterResumeEvents([In, MarshalAs(UnmanagedType.IDispatch)] object app)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterResumeEvents");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(app);
				return;
			}

			NetOffice.VisioApi.IVApplication newapp = Factory.CreateObjectFromComProxy(_eventClass, app) as NetOffice.VisioApi.IVApplication;
			object[] paramsArray = new object[1];
			paramsArray[0] = newapp;
			_eventBinding.RaiseCustomEvent("AfterResumeEvents", ref paramsArray);
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

		public void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDataRecordsetDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(dataRecordset);
				return;
			}

			NetOffice.VisioApi.IVDataRecordset newDataRecordset = Factory.CreateObjectFromComProxy(_eventClass, dataRecordset) as NetOffice.VisioApi.IVDataRecordset;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordset;
			_eventBinding.RaiseCustomEvent("BeforeDataRecordsetDelete", ref paramsArray);
		}

		public void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DataRecordsetChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(dataRecordsetChanged);
				return;
			}

			NetOffice.VisioApi.IVDataRecordsetChangedEvent newDataRecordsetChanged = Factory.CreateObjectFromComProxy(_eventClass, dataRecordsetChanged) as NetOffice.VisioApi.IVDataRecordsetChangedEvent;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordsetChanged;
			_eventBinding.RaiseCustomEvent("DataRecordsetChanged", ref paramsArray);
		}

		public void DataRecordsetAdded([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DataRecordsetAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(dataRecordset);
				return;
			}

			NetOffice.VisioApi.IVDataRecordset newDataRecordset = Factory.CreateObjectFromComProxy(_eventClass, dataRecordset) as NetOffice.VisioApi.IVDataRecordset;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordset;
			_eventBinding.RaiseCustomEvent("DataRecordsetAdded", ref paramsArray);
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

		public void AfterRemoveHiddenInformation([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterRemoveHiddenInformation");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.VisioApi.IVDocument newdoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.VisioApi.IVDocument;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdoc;
			_eventBinding.RaiseCustomEvent("AfterRemoveHiddenInformation", ref paramsArray);
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

		public void RuleSetValidated([In, MarshalAs(UnmanagedType.IDispatch)] object ruleSet)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("RuleSetValidated");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(ruleSet);
				return;
			}

			NetOffice.VisioApi.IVValidationRuleSet newRuleSet = Factory.CreateObjectFromComProxy(_eventClass, ruleSet) as NetOffice.VisioApi.IVValidationRuleSet;
			object[] paramsArray = new object[1];
			paramsArray[0] = newRuleSet;
			_eventBinding.RaiseCustomEvent("RuleSetValidated", ref paramsArray);
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