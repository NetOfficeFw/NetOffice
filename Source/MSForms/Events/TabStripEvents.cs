using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi.EventContracts
{
    /// <summary>
    /// TabStripEvents
    /// </summary>
	[SupportByVersion("MSForms", 2)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("7B020EC7-AF6C-11CE-9F46-00AA00574A4F"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface TabStripEvents
	{
        /// <summary>
        /// BeforeDragOver
        /// </summary>
        /// <param name="index"></param>
        /// <param name="cancel"></param>
        /// <param name="data"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="dragState"></param>
        /// <param name="effect"></param>
        /// <param name="shift"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [SinkArgument("data", typeof(MSFormsApi.DataObject))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [SinkArgument("dragState", SinkArgumentType.Enum, typeof(MSFormsApi.Enums.fmDragState))]
        [SinkArgument("effect", typeof(MSFormsApi.ReturnEffect))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void BeforeDragOver([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In] object dragState, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

        /// <summary>
        /// BeforeDropOrPaste
        /// </summary>
        /// <param name="index"></param>
        /// <param name="cancel"></param>
        /// <param name="action"></param>
        /// <param name="data"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="effect"></param>
        /// <param name="shift"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [SinkArgument("action", SinkArgumentType.Enum, typeof(MSFormsApi.Enums.fmAction))]
        [SinkArgument("data", typeof(MSFormsApi.DataObject))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [SinkArgument("effect", typeof(MSFormsApi.ReturnEffect))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void BeforeDropOrPaste([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

        /// <summary>
        /// Change
        /// </summary>
		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void Change();

        /// <summary>
        /// Click
        /// </summary>
        /// <param name="index"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click([In] object index);

        /// <summary>
        /// DblClick
        /// </summary>
        /// <param name="index"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

        /// <summary>
        /// Error
        /// </summary>
        /// <param name="number"></param>
        /// <param name="description"></param>
        /// <param name="sCode"></param>
        /// <param name="source"></param>
        /// <param name="helpFile"></param>
        /// <param name="helpContext"></param>
        /// <param name="cancelDisplay"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("number", SinkArgumentType.Int32)]
        [SinkArgument("description", typeof(MSFormsApi.ReturnString))]
        [SinkArgument("sCode", SinkArgumentType.Int32)]
        [SinkArgument("source", SinkArgumentType.String)]
        [SinkArgument("helpFile", SinkArgumentType.String)]
        [SinkArgument("helpContext", SinkArgumentType.Int32)]
        [SinkArgument("cancelDisplay", typeof(MSFormsApi.ReturnBoolean))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-608)]
		void Error([In] object number, [In, MarshalAs(UnmanagedType.IDispatch)] object description, [In] object sCode, [In] object source, [In] object helpFile, [In] object helpContext, [In, MarshalAs(UnmanagedType.IDispatch)] object cancelDisplay);

        /// <summary>
        /// KeyDown
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("keyCode", typeof(MSFormsApi.ReturnInteger))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift);

        /// <summary>
        /// KeyPress
        /// </summary>
        /// <param name="keyAscii"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("keyAscii", typeof(MSFormsApi.ReturnInteger))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In, MarshalAs(UnmanagedType.IDispatch)] object keyAscii);

        /// <summary>
        /// KeyUp
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("keyCode", typeof(MSFormsApi.ReturnInteger))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift);

        /// <summary>
        /// MouseDown
        /// </summary>
        /// <param name="index"></param>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseMove
        /// </summary>
        /// <param name="index"></param>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseUp
        /// </summary>
        /// <param name="index"></param>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);
	}
}
