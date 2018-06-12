using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.WordApi.EventContracts
{
    /// <summary>
    /// ApplicationEvents2
    /// </summary>
    [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000209FE-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface ApplicationEvents2
    {
        /// <summary>
        /// Startup
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        void Startup();

        /// <summary>
        /// Quit
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
        void Quit();

        /// <summary>
        /// DocumentChange
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
        void DocumentChange();

        /// <summary>
        /// DocumentOpen
        /// </summary>
        /// <param name="doc"></param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
        void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

        /// <summary>
        /// DocumentBeforeClose
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
        void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

        /// <summary>
        /// DocumentBeforePrint
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
        void DocumentBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

        /// <summary>
        /// DocumentBeforeSave
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="saveAsUI"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("saveAsUI", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
        void DocumentBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object saveAsUI, [In] [Out] ref object cancel);

        /// <summary>
        /// NewDocument
        /// </summary>
        /// <param name="doc"></param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
        void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

        /// <summary>
        /// WindowActivate
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="wn"></param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("wn", typeof(WordApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
        void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// WindowDeactivate
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="wn"></param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("wn", typeof(WordApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
        void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// WindowSelectionChange
        /// </summary>
        /// <param name="sel"></param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
        void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
        void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

        /// <summary>
        /// WindowBeforeDoubleClick
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="cancel"></param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
        void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);
    }
}
