using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi.EventContracts
{
    /// <summary>
    /// EApplication
    /// </summary>
    [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("914934C2-5A91-11CF-8700-00AA0060263B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EApplication
	{
        /// <summary>
        /// WindowSelectionChange
        /// </summary>
        /// <param name="sel"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("sel", typeof(PowerPointApi.Selection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2001)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel);

        /// <summary>
        /// WindowBeforeRightClick
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("sel", typeof(PowerPointApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2002)]
		void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

        /// <summary>
        /// WindowBeforeDoubleClick
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("sel", typeof(PowerPointApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2003)]
		void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

        /// <summary>
        /// PresentationClose
        /// </summary>
        /// <param name="pres"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2004)]
		void PresentationClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

        /// <summary>
        /// PresentationSave
        /// </summary>
        /// <param name="pres"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2005)]
		void PresentationSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

        /// <summary>
        /// PresentationOpen
        /// </summary>
        /// <param name="pres"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2006)]
		void PresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

        /// <summary>
        /// NewPresentation
        /// </summary>
        /// <param name="pres"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2007)]
		void NewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

        /// <summary>
        /// PresentationNewSlide
        /// </summary>
        /// <param name="sld"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("sld", typeof(PowerPointApi.Slide))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2008)]
		void PresentationNewSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld);

        /// <summary>
        /// WindowActivate
        /// </summary>
        /// <param name="pres"></param>
        /// <param name="wn"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [SinkArgument("wn", typeof(PowerPointApi.DocumentWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2009)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// WindowDeactivate
        /// </summary>
        /// <param name="pres"></param>
        /// <param name="wn"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [SinkArgument("wn", typeof(PowerPointApi.DocumentWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2010)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// SlideShowBegin
        /// </summary>
        /// <param name="wn"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.DocumentWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2011)]
		void SlideShowBegin([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// SlideShowNextBuild
        /// </summary>
        /// <param name="wn"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.DocumentWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2012)]
		void SlideShowNextBuild([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// SlideShowNextSlide
        /// </summary>
        /// <param name="wn"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.DocumentWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2013)]
		void SlideShowNextSlide([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// SlideShowEnd
        /// </summary>
        /// <param name="pres"></param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2014)]
		void SlideShowEnd([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

        /// <summary>
        /// PresentationPrint
        /// </summary>
        /// <param name="pres"></param>
        [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2015)]
		void PresentationPrint([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

        /// <summary>
        /// SlideSelectionChanged
        /// </summary>
        /// <param name="sldRange"></param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
        [SinkArgument("sldRange", typeof(PowerPointApi.SlideRange))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2016)]
		void SlideSelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange);

        /// <summary>
        /// ColorSchemeChanged
        /// </summary>
        /// <param name="sldRange"></param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
        [SinkArgument("sldRange", typeof(PowerPointApi.SlideRange))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2017)]
		void ColorSchemeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange);

        /// <summary>
        /// PresentationBeforeSave
        /// </summary>
        /// <param name="pres"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2018)]
		void PresentationBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel);

        /// <summary>
        /// SlideShowNextClick
        /// </summary>
        /// <param name="wn"></param>
        /// <param name="nEffect"></param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.SlideShowWindow))]
        [SinkArgument("nEffect", typeof(PowerPointApi.Effect))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2019)]
		void SlideShowNextClick([In, MarshalAs(UnmanagedType.IDispatch)] object wn, [In, MarshalAs(UnmanagedType.IDispatch)] object nEffect);

        /// <summary>
        /// AfterNewPresentation
        /// </summary>
        /// <param name="pres"></param>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2020)]
		void AfterNewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

        /// <summary>
        /// AfterPresentationOpen
        /// </summary>
        /// <param name="pres"></param>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2021)]
		void AfterPresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

        /// <summary>
        /// PresentationSync
        /// </summary>
        /// <param name="pres"></param>
        /// <param name="syncEventType"></param>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [SinkArgument("syncEventType", SinkArgumentType.Enum, typeof(OfficeApi.Enums.MsoSyncEventType))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2022)]
		void PresentationSync([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] object syncEventType);

        /// <summary>
        /// SlideShowOnNext
        /// </summary>
        /// <param name="wn"></param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.SlideShowWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2023)]
		void SlideShowOnNext([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// SlideShowOnPrevious
        /// </summary>
        /// <param name="wn"></param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.SlideShowWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2024)]
		void SlideShowOnPrevious([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// PresentationBeforeClose
        /// </summary>
        /// <param name="pres"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2025)]
		void PresentationBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel);

        /// <summary>
        /// ProtectedViewWindowOpen
        /// </summary>
        /// <param name="protViewWindow"></param>
		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("protViewWindow", typeof(PowerPointApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2026)]
		void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow);

        /// <summary>
        /// ProtectedViewWindowBeforeEdit
        /// </summary>
        /// <param name="protViewWindow"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("protViewWindow", typeof(PowerPointApi.ProtectedViewWindow))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2027)]
		void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] [Out] ref object cancel);

        /// <summary>
        /// ProtectedViewWindowBeforeClose
        /// </summary>
        /// <param name="protViewWindow"></param>
        /// <param name="protectedViewCloseReason"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("protViewWindow", typeof(PowerPointApi.ProtectedViewWindow))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2028)]
		void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] object protectedViewCloseReason, [In] [Out] ref object cancel);

        /// <summary>
        /// ProtectedViewWindowActivate
        /// </summary>
        /// <param name="protViewWindow"></param>
		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("protViewWindow", typeof(PowerPointApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2029)]
		void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow);

        /// <summary>
        /// ProtectedViewWindowDeactivate
        /// </summary>
        /// <param name="protViewWindow"></param>
		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("protViewWindow", typeof(PowerPointApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2030)]
		void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow);

        /// <summary>
        /// PresentationCloseFinal
        /// </summary>
        /// <param name="pres"></param>
		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2031)]
		void PresentationCloseFinal([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

        /// <summary>
        /// AfterDragDropOnSlide
        /// </summary>
        /// <param name="sld"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
		[SupportByVersion("PowerPoint", 15, 16)]
        [SinkArgument("sld", typeof(PowerPointApi.Slide))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]    
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2032)]
		void AfterDragDropOnSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld, [In] object x, [In] object y);

        /// <summary>
        /// AfterShapeSizeChange
        /// </summary>
        /// <param name="shp"></param>
		[SupportByVersion("PowerPoint", 15, 16)]
        [SinkArgument("shp", typeof(PowerPointApi.Shape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2033)]
		void AfterShapeSizeChange([In, MarshalAs(UnmanagedType.IDispatch)] object shp);
	}
}
