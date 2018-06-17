using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi.EventContracts
{
    /// <summary>
    /// FormEvents
    /// </summary>
	[SupportByVersion("MSForms", 2)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("5B9D8FC8-4A71-101B-97A6-00000B65C08B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface FormEvents
	{
        /// <summary>
        /// AddControl
        /// </summary>
        /// <param name="control"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("control", typeof(MSFormsApi.Control))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(768)]
		void AddControl([In, MarshalAs(UnmanagedType.IDispatch)] object control);

        /// <summary>
        /// BeforeDragOver
        /// </summary>
        /// <param name="cancel"></param>
        /// <param name="control"></param>
        /// <param name="data"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="state"></param>
        /// <param name="effect"></param>
        /// <param name="shift"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [SinkArgument("control", typeof(MSFormsApi.Control))]
        [SinkArgument("data", typeof(MSFormsApi.DataObject))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [SinkArgument("state", SinkArgumentType.Enum, typeof(MSFormsApi.Enums.fmDragState))]
        [SinkArgument("effect", typeof(MSFormsApi.ReturnEffect))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void BeforeDragOver([In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object control, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In] object state, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

        /// <summary>
        /// BeforeDropOrPaste
        /// </summary>
        /// <param name="cancel"></param>
        /// <param name="control"></param>
        /// <param name="action"></param>
        /// <param name="data"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="effect"></param>
        /// <param name="shift"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [SinkArgument("control", typeof(MSFormsApi.Control))]
        [SinkArgument("action", typeof(MSFormsApi.Enums.fmAction))]
        [SinkArgument("data", typeof(MSFormsApi.DataObject))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [SinkArgument("effect", typeof(MSFormsApi.ReturnEffect))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void BeforeDropOrPaste([In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object control, [In] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

        /// <summary>
        /// Click
        /// </summary>
		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

        /// <summary>
        /// DblClick
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

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
        [SinkArgument("number", SinkArgumentType.Int16)]
        [SinkArgument("description", typeof(MSFormsApi.ReturnString))]
        [SinkArgument("sCode", SinkArgumentType.Int16)]
        [SinkArgument("source", SinkArgumentType.String)]
        [SinkArgument("helpFile", SinkArgumentType.String)]
        [SinkArgument("helpContext", SinkArgumentType.Int16)]
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
        /// Layout
        /// </summary>
		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(770)]
		void Layout();

        /// <summary>
        /// MouseDown
        /// </summary>
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
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseMove
        /// </summary>
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
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseUp
        /// </summary>
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
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// RemoveControl
        /// </summary>
        /// <param name="control"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("control", typeof(MSFormsApi.Control))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(771)]
		void RemoveControl([In, MarshalAs(UnmanagedType.IDispatch)] object control);

        /// <summary>
        /// Scroll
        /// </summary>
        /// <param name="actionX"></param>
        /// <param name="actionY"></param>
        /// <param name="requestDx"></param>
        /// <param name="requestDy"></param>
        /// <param name="actualDx"></param>
        /// <param name="actualDy"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("actionX", SinkArgumentType.Enum, typeof(MSFormsApi.Enums.fmScrollAction))]
        [SinkArgument("actionY", SinkArgumentType.Enum, typeof(MSFormsApi.Enums.fmScrollAction))]
        [SinkArgument("requestDx", SinkArgumentType.Single)]
        [SinkArgument("requestDy", SinkArgumentType.Single)]
        [SinkArgument("actualDx", typeof(MSFormsApi.ReturnSingle))]
        [SinkArgument("actualDy", typeof(MSFormsApi.ReturnSingle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(772)]
		void Scroll([In] object actionX, [In] object actionY, [In] object requestDx, [In] object requestDy, [In, MarshalAs(UnmanagedType.IDispatch)] object actualDx, [In, MarshalAs(UnmanagedType.IDispatch)] object actualDy);

        /// <summary>
        /// Zoom
        /// </summary>
        /// <param name="percent"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("percent", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(773)]
		void Zoom([In] [Out] ref object percent);
	}
}
