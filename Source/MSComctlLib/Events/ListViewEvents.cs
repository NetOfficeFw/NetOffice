using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi.EventContracts
{
    /// <summary>
    /// ListViewEvents
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("BDD1F04A-858B-11D1-B16A-00C0F0283628"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ListViewEvents
	{
        /// <summary>
        /// BeforeLabelEdit
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void BeforeLabelEdit([In] [Out] ref object cancel);

        /// <summary>
        /// AfterLabelEdit
        /// </summary>
        /// <param name="cancel"></param>
        /// <param name="newString"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [SinkArgument("newString", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void AfterLabelEdit([In] [Out] ref object cancel, [In] [Out] ref object newString);

        /// <summary>
        /// ColumnClick
        /// </summary>
        /// <param name="columnHeader"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("columnHeader", typeof(MSComctlLibApi.ColumnHeader))]        
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void ColumnClick([In, MarshalAs(UnmanagedType.IDispatch)] object columnHeader);

        /// <summary>
        /// ItemClick
        /// </summary>
        /// <param name="item"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("columnHeader", typeof(MSComctlLibApi.ListItem))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void ItemClick([In, MarshalAs(UnmanagedType.IDispatch)] object item);

        /// <summary>
        /// KeyDown
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("keyCode", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In] [Out] ref object keyCode, [In] object shift);

        /// <summary>
        /// KeyUp
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("keyCode", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In] [Out] ref object keyCode, [In] object shift);

        /// <summary>
        /// KeyPress
        /// </summary>
        /// <param name="keyAscii"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("keyAscii", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In] [Out] ref object keyAscii);

        /// <summary>
        /// MouseDown
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseMove
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// MouseUp
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("button", SinkArgumentType.Int32)]
        [SinkArgument("shift", SinkArgumentType.Int32)]
        [SinkArgument("x", SinkArgumentType.Int32)]
        [SinkArgument("y", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

        /// <summary>
        /// Click
        /// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

        /// <summary>
        /// DblClick
        /// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick();

        /// <summary>
        /// OLEStartDrag
        /// </summary>
        /// <param name="data"></param>
        /// <param name="allowedEffects"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("data", typeof(MSComctlLibApi.DataObject))]
        [SinkArgument("allowedEffects", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1550)]
		void OLEStartDrag([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object allowedEffects);

        /// <summary>
        /// OLEGiveFeedback
        /// </summary>
        /// <param name="effect"></param>
        /// <param name="defaultCursors"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("effect", SinkArgumentType.Int32)]
        [SinkArgument("defaultCursors", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1551)]
		void OLEGiveFeedback([In] [Out] ref object effect, [In] [Out] ref object defaultCursors);

        /// <summary>
        /// OLESetData
        /// </summary>
        /// <param name="data"></param>
        /// <param name="dataFormat"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("data", typeof(MSComctlLibApi.DataObject))]
        [SinkArgument("dataFormat", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1552)]
		void OLESetData([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object dataFormat);

        /// <summary>
        /// OLECompleteDrag
        /// </summary>
        /// <param name="effect"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("effect", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1553)]
		void OLECompleteDrag([In] [Out] ref object effect);

        /// <summary>
        /// OLEDragOver
        /// </summary>
        /// <param name="data"></param>
        /// <param name="effect"></param>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="state"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("data", typeof(MSComctlLibApi.DataObject))]
        [SinkArgument("effect", SinkArgumentType.Int32)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [SinkArgument("state", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1554)]
		void OLEDragOver([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object effect, [In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y, [In] [Out] ref object state);

        /// <summary>
        /// OLEDragDrop
        /// </summary>
        /// <param name="data"></param>
        /// <param name="effect"></param>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("data", typeof(MSComctlLibApi.DataObject))]
        [SinkArgument("effect", SinkArgumentType.Int32)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1555)]
		void OLEDragDrop([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object effect, [In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

        /// <summary>
        /// ItemCheck
        /// </summary>
        /// <param name="item"></param>
		[SupportByVersion("MSComctlLib", 6)]
        [SinkArgument("columnHeader", typeof(MSComctlLibApi.ListItem))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void ItemCheck([In, MarshalAs(UnmanagedType.IDispatch)] object item);
	}
}
