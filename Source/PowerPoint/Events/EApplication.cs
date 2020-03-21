using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("914934C2-5A91-11CF-8700-00AA0060263B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EApplication
	{
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("sel", typeof(PowerPointApi.Selection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2001)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("sel", typeof(PowerPointApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2002)]
		void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("sel", typeof(PowerPointApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2003)]
		void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2004)]
		void PresentationClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2005)]
		void PresentationSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2006)]
		void PresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2007)]
		void NewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("sld", typeof(PowerPointApi.Slide))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2008)]
		void PresentationNewSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [SinkArgument("wn", typeof(PowerPointApi.DocumentWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2009)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [SinkArgument("wn", typeof(PowerPointApi.DocumentWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2010)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.SlideShowWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2011)]
		void SlideShowBegin([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.SlideShowWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2012)]
		void SlideShowNextBuild([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.SlideShowWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2013)]
		void SlideShowNextSlide([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2014)]
		void SlideShowEnd([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

        [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2015)]
		void PresentationPrint([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
        [SinkArgument("sldRange", typeof(PowerPointApi.SlideRange))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2016)]
		void SlideSelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange);

		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
        [SinkArgument("sldRange", typeof(PowerPointApi.SlideRange))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2017)]
		void ColorSchemeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange);

		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2018)]
		void PresentationBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel);

		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.SlideShowWindow))]
        [SinkArgument("nEffect", typeof(PowerPointApi.Effect))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2019)]
		void SlideShowNextClick([In, MarshalAs(UnmanagedType.IDispatch)] object wn, [In, MarshalAs(UnmanagedType.IDispatch)] object nEffect);

		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2020)]
		void AfterNewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2021)]
		void AfterPresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [SinkArgument("syncEventType", SinkArgumentType.Enum, typeof(OfficeApi.Enums.MsoSyncEventType))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2022)]
		void PresentationSync([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] object syncEventType);

		[SupportByVersion("PowerPoint", 12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.SlideShowWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2023)]
		void SlideShowOnNext([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("PowerPoint", 12,14,15,16)]
        [SinkArgument("wn", typeof(PowerPointApi.SlideShowWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2024)]
		void SlideShowOnPrevious([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2025)]
		void PresentationBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel);

		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("protViewWindow", typeof(PowerPointApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2026)]
		void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow);

		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("protViewWindow", typeof(PowerPointApi.ProtectedViewWindow))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2027)]
		void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] [Out] ref object cancel);

		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("protViewWindow", typeof(PowerPointApi.ProtectedViewWindow))]
        [SinkArgument("protectedViewCloseReason", SinkArgumentType.Enum , typeof(PowerPointApi.Enums.PpProtectedViewCloseReason))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2028)]
		void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] object protectedViewCloseReason, [In] [Out] ref object cancel);

		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("protViewWindow", typeof(PowerPointApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2029)]
		void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow);

		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("protViewWindow", typeof(PowerPointApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2030)]
		void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow);

		[SupportByVersion("PowerPoint", 14,15,16)]
        [SinkArgument("pres", typeof(PowerPointApi.Presentation))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2031)]
		void PresentationCloseFinal([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersion("PowerPoint", 15, 16)]
        [SinkArgument("sld", typeof(PowerPointApi.Slide))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]    
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2032)]
		void AfterDragDropOnSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld, [In] object x, [In] object y);

		[SupportByVersion("PowerPoint", 15, 16)]
        [SinkArgument("shp", typeof(PowerPointApi.Shape))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2033)]
		void AfterShapeSizeChange([In, MarshalAs(UnmanagedType.IDispatch)] object shp);
	}
	
	#endregion
	
	#region SinkHelper
	
    [InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EApplication_SinkHelper : SinkHelper, EApplication
	{
		#region Static
		
		public static readonly string Id = "914934C2-5A91-11CF-8700-00AA0060263B";
		
		#endregion
	
		#region Ctor

		public EApplication_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region EApplication
		
        public void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel)
		{
            if (!Validate("WindowSelectionChange"))
            {
                Invoker.ReleaseParamsArray(sel);
                return;
            }

			NetOffice.PowerPointApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Selection>(EventClass, sel, NetOffice.PowerPointApi.Selection.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newSel;
			EventBinding.RaiseCustomEvent("WindowSelectionChange", ref paramsArray);
		}

        public void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
		{
            if (!Validate("WindowBeforeRightClick"))
            {
                Invoker.ReleaseParamsArray(sel, cancel);
                return;
            }

            NetOffice.PowerPointApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Selection>(EventClass, sel, NetOffice.PowerPointApi.Selection.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newSel;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("WindowBeforeRightClick", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        public void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
        {
            if (!Validate("WindowBeforeDoubleClick"))
            {
                Invoker.ReleaseParamsArray(sel, cancel);
                return;
            }

            NetOffice.PowerPointApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Selection>(EventClass, sel, NetOffice.PowerPointApi.Selection.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newSel;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("WindowBeforeDoubleClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void PresentationClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("PresentationClose"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			EventBinding.RaiseCustomEvent("PresentationClose", ref paramsArray);
		}

        public void PresentationSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
            if (!Validate("PresentationSave"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			EventBinding.RaiseCustomEvent("PresentationSave", ref paramsArray);
		}

        public void PresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
            if (!Validate("PresentationOpen"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			EventBinding.RaiseCustomEvent("PresentationOpen", ref paramsArray);
		}

        public void NewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
            if (!Validate("NewPresentation"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			EventBinding.RaiseCustomEvent("NewPresentation", ref paramsArray);
		}

        public void PresentationNewSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld)
        {
            if (!Validate("PresentationNewSlide"))
            {
                Invoker.ReleaseParamsArray(sld);
                return;
            }

			NetOffice.PowerPointApi.Slide newSld = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Slide>(EventClass, sld, NetOffice.PowerPointApi.Slide.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newSld;
			EventBinding.RaiseCustomEvent("PresentationNewSlide", ref paramsArray);
		}

        public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowActivate"))
            {
                Invoker.ReleaseParamsArray(pres, wn);
                return;
            }

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
			NetOffice.PowerPointApi.DocumentWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.DocumentWindow>(EventClass, wn, NetOffice.PowerPointApi.DocumentWindow.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newPres;
			paramsArray[1] = newWn;
			EventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
		}

        public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
            if (!Validate("WindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(pres, wn);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
            NetOffice.PowerPointApi.DocumentWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.DocumentWindow>(EventClass, wn, NetOffice.PowerPointApi.DocumentWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newPres;
			paramsArray[1] = newWn;
			EventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
		}

        public void SlideShowBegin([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("SlideShowBegin"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideShowWindow>(EventClass, wn, NetOffice.PowerPointApi.DocumentWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			EventBinding.RaiseCustomEvent("SlideShowBegin", ref paramsArray);
		}

        public void SlideShowNextBuild([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("SlideShowNextBuild"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideShowWindow>(EventClass, wn, NetOffice.PowerPointApi.DocumentWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			EventBinding.RaiseCustomEvent("SlideShowNextBuild", ref paramsArray);
		}

        public void SlideShowNextSlide([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("SlideShowNextSlide"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideShowWindow>(EventClass, wn, NetOffice.PowerPointApi.DocumentWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			EventBinding.RaiseCustomEvent("SlideShowNextSlide", ref paramsArray);
		}

        public void SlideShowEnd([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
            if (!Validate("SlideShowEnd"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			EventBinding.RaiseCustomEvent("SlideShowEnd", ref paramsArray);
		}

        public void PresentationPrint([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
            if (!Validate("PresentationPrint"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			EventBinding.RaiseCustomEvent("PresentationPrint", ref paramsArray);
		}

        public void SlideSelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange)
        {
            if (!Validate("SlideSelectionChanged"))
            {
                Invoker.ReleaseParamsArray(sldRange);
                return;
            }

			NetOffice.PowerPointApi.SlideRange newSldRange = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideRange>(EventClass, sldRange, NetOffice.PowerPointApi.SlideRange.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newSldRange;
			EventBinding.RaiseCustomEvent("SlideSelectionChanged", ref paramsArray);
		}

        public void ColorSchemeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange)
		{
            if (!Validate("ColorSchemeChanged"))
            {
                Invoker.ReleaseParamsArray(sldRange);
                return;
            }

            NetOffice.PowerPointApi.SlideRange newSldRange = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideRange>(EventClass, sldRange, NetOffice.PowerPointApi.SlideRange.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newSldRange;
			EventBinding.RaiseCustomEvent("ColorSchemeChanged", ref paramsArray);
		}

        public void PresentationBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel)
        {
            if (!Validate("PresentationBeforeSave"))
            {
                Invoker.ReleaseParamsArray(pres, cancel);
                return;
            }

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newPres;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("PresentationBeforeSave", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        public void SlideShowNextClick([In, MarshalAs(UnmanagedType.IDispatch)] object wn, [In, MarshalAs(UnmanagedType.IDispatch)] object nEffect)
        {
            if (!Validate("SlideShowNextClick"))
            {
                Invoker.ReleaseParamsArray(wn, nEffect);
                return;
            }

            NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideShowWindow>(EventClass, wn, NetOffice.PowerPointApi.SlideShowWindow.LateBindingApiWrapperType);
			NetOffice.PowerPointApi.Effect newnEffect = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Effect>(EventClass, nEffect, NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newWn;
			paramsArray[1] = newnEffect;
			EventBinding.RaiseCustomEvent("SlideShowNextClick", ref paramsArray);
		}

        public void AfterNewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("AfterNewPresentation"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			EventBinding.RaiseCustomEvent("AfterNewPresentation", ref paramsArray);
		}

        public void AfterPresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("AfterPresentationOpen"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			EventBinding.RaiseCustomEvent("AfterPresentationOpen", ref paramsArray);
		}

        public void PresentationSync([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] object syncEventType)
        {
            if (!Validate("PresentationSync"))
            {
                Invoker.ReleaseParamsArray(pres, syncEventType);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
            NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
			object[] paramsArray = new object[2];
			paramsArray[0] = newPres;
			paramsArray[1] = newSyncEventType;
			EventBinding.RaiseCustomEvent("PresentationSync", ref paramsArray);
		}

        public void SlideShowOnNext([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("SlideShowOnNext"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

			NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideShowWindow>(EventClass, wn, NetOffice.PowerPointApi.SlideShowWindow.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			EventBinding.RaiseCustomEvent("SlideShowOnNext", ref paramsArray);
		}

        public void SlideShowOnPrevious([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("SlideShowOnPrevious"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideShowWindow>(EventClass, wn, NetOffice.PowerPointApi.SlideShowWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			EventBinding.RaiseCustomEvent("SlideShowOnPrevious", ref paramsArray);
		}

        public void PresentationBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel)
        {
            if (!Validate("SlideShowOnPrevious"))
            {
                Invoker.ReleaseParamsArray(pres, cancel);
                return;
            }

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newPres;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("PresentationBeforeClose", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        public void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow)
        {
            if (!Validate("ProtectedViewWindowOpen"))
            {
                Invoker.ReleaseParamsArray(protViewWindow);
                return;
            }

			NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.ProtectedViewWindow>(EventClass, protViewWindow, NetOffice.PowerPointApi.ProtectedViewWindow.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newProtViewWindow;
			EventBinding.RaiseCustomEvent("ProtectedViewWindowOpen", ref paramsArray);
		}
        
        public void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] [Out] ref object cancel)
        {
            if (!Validate("ProtectedViewWindowBeforeEdit"))
            {
                Invoker.ReleaseParamsArray(protViewWindow, cancel);
                return;
            }

            NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.ProtectedViewWindow>(EventClass, protViewWindow, NetOffice.PowerPointApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newProtViewWindow;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeEdit", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        public void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] object protectedViewCloseReason, [In] [Out] ref object cancel)
        {
            if (!Validate("ProtectedViewWindowBeforeClose"))
            {
                Invoker.ReleaseParamsArray(protViewWindow, protectedViewCloseReason, cancel);
                return;
            }

            NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.ProtectedViewWindow>(EventClass, protViewWindow, NetOffice.PowerPointApi.ProtectedViewWindow.LateBindingApiWrapperType);
            NetOffice.PowerPointApi.Enums.PpProtectedViewCloseReason newProtectedViewCloseReason = (NetOffice.PowerPointApi.Enums.PpProtectedViewCloseReason)protectedViewCloseReason;
			object[] paramsArray = new object[3];
			paramsArray[0] = newProtViewWindow;
			paramsArray[1] = newProtectedViewCloseReason;
			paramsArray.SetValue(cancel, 2);
			EventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeClose", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        public void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow)
        {
            if (!Validate("ProtectedViewWindowActivate"))
            {
                Invoker.ReleaseParamsArray(protViewWindow);
                return;
            }

            NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.ProtectedViewWindow>(EventClass, protViewWindow, NetOffice.PowerPointApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newProtViewWindow;
			EventBinding.RaiseCustomEvent("ProtectedViewWindowActivate", ref paramsArray);
		}

        public void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow)
        {
            if (!Validate("ProtectedViewWindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(protViewWindow);
                return;
            }

            NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.ProtectedViewWindow>(EventClass, protViewWindow, NetOffice.PowerPointApi.ProtectedViewWindow.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newProtViewWindow;
			EventBinding.RaiseCustomEvent("ProtectedViewWindowDeactivate", ref paramsArray);
		}

        public void PresentationCloseFinal([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("PresentationCloseFinal"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			EventBinding.RaiseCustomEvent("PresentationCloseFinal", ref paramsArray);
		}

        public void AfterDragDropOnSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld, [In] object x, [In] object y)
        {
            if (!Validate("AfterDragDropOnSlide"))
            {
                Invoker.ReleaseParamsArray(sld, x, y);
                return;
            }

			NetOffice.PowerPointApi.Slide newSld = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Slide>(EventClass, sld, NetOffice.PowerPointApi.Slide.LateBindingApiWrapperType);
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[3];
			paramsArray[0] = newSld;
			paramsArray[1] = newX;
			paramsArray[2] = newY;
			EventBinding.RaiseCustomEvent("AfterDragDropOnSlide", ref paramsArray);
		}

        public void AfterShapeSizeChange([In, MarshalAs(UnmanagedType.IDispatch)] object shp)
		{
            if (!Validate("AfterShapeSizeChange"))
            {
                Invoker.ReleaseParamsArray(shp);
                return;
            }

			NetOffice.PowerPointApi.Shape newshp = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Shape>(EventClass, shp, NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newshp;
			EventBinding.RaiseCustomEvent("AfterShapeSizeChange", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}