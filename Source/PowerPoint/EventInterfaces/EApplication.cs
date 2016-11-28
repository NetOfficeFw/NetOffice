using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.PowerPointApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
	[ComImport, Guid("914934C2-5A91-11CF-8700-00AA0060263B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EApplication
	{
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2001)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2002)]
		void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2003)]
		void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2004)]
		void PresentationClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2005)]
		void PresentationSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2006)]
		void PresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2007)]
		void NewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2008)]
		void PresentationNewSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2009)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2010)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2011)]
		void SlideShowBegin([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2012)]
		void SlideShowNextBuild([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2013)]
		void SlideShowNextSlide([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2014)]
		void SlideShowEnd([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2015)]
		void PresentationPrint([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2016)]
		void SlideSelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange);

		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2017)]
		void ColorSchemeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange);

		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2018)]
		void PresentationBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2019)]
		void SlideShowNextClick([In, MarshalAs(UnmanagedType.IDispatch)] object wn, [In, MarshalAs(UnmanagedType.IDispatch)] object nEffect);

		[SupportByVersionAttribute("PowerPoint", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2020)]
		void AfterNewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersionAttribute("PowerPoint", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2021)]
		void AfterPresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersionAttribute("PowerPoint", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2022)]
		void PresentationSync([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] object syncEventType);

		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2023)]
		void SlideShowOnNext([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2024)]
		void SlideShowOnPrevious([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2025)]
		void PresentationBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2026)]
		void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow);

		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2027)]
		void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2028)]
		void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] object protectedViewCloseReason, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2029)]
		void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow);

		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2030)]
		void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow);

		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2031)]
		void PresentationCloseFinal([In, MarshalAs(UnmanagedType.IDispatch)] object pres);

		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2032)]
		void AfterDragDropOnSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld, [In] object x, [In] object y);

		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2033)]
		void AfterShapeSizeChange([In, MarshalAs(UnmanagedType.IDispatch)] object shp);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EApplication_SinkHelper : SinkHelper, EApplication
	{
		#region Static
		
		public static readonly string Id = "914934C2-5A91-11CF-8700-00AA0060263B";
		
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
		
		public void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowSelectionChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sel);
				return;
			}

			NetOffice.PowerPointApi.Selection newSel = Factory.CreateObjectFromComProxy(_eventClass, sel) as NetOffice.PowerPointApi.Selection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSel;
			_eventBinding.RaiseCustomEvent("WindowSelectionChange", ref paramsArray);
		}

		public void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowBeforeRightClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sel, cancel);
				return;
			}

			NetOffice.PowerPointApi.Selection newSel = Factory.CreateObjectFromComProxy(_eventClass, sel) as NetOffice.PowerPointApi.Selection;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSel;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("WindowBeforeRightClick", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowBeforeDoubleClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sel, cancel);
				return;
			}

			NetOffice.PowerPointApi.Selection newSel = Factory.CreateObjectFromComProxy(_eventClass, sel) as NetOffice.PowerPointApi.Selection;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSel;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("WindowBeforeDoubleClick", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void PresentationClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PresentationClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			_eventBinding.RaiseCustomEvent("PresentationClose", ref paramsArray);
		}

		public void PresentationSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PresentationSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			_eventBinding.RaiseCustomEvent("PresentationSave", ref paramsArray);
		}

		public void PresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PresentationOpen");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			_eventBinding.RaiseCustomEvent("PresentationOpen", ref paramsArray);
		}

		public void NewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewPresentation");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			_eventBinding.RaiseCustomEvent("NewPresentation", ref paramsArray);
		}

		public void PresentationNewSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PresentationNewSlide");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sld);
				return;
			}

			NetOffice.PowerPointApi.Slide newSld = Factory.CreateObjectFromComProxy(_eventClass, sld) as NetOffice.PowerPointApi.Slide;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSld;
			_eventBinding.RaiseCustomEvent("PresentationNewSlide", ref paramsArray);
		}

		public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres, wn);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			NetOffice.PowerPointApi.DocumentWindow newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.PowerPointApi.DocumentWindow;
			object[] paramsArray = new object[2];
			paramsArray[0] = newPres;
			paramsArray[1] = newWn;
			_eventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
		}

		public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres, wn);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			NetOffice.PowerPointApi.DocumentWindow newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.PowerPointApi.DocumentWindow;
			object[] paramsArray = new object[2];
			paramsArray[0] = newPres;
			paramsArray[1] = newWn;
			_eventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
		}

		public void SlideShowBegin([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SlideShowBegin");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.PowerPointApi.SlideShowWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			_eventBinding.RaiseCustomEvent("SlideShowBegin", ref paramsArray);
		}

		public void SlideShowNextBuild([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SlideShowNextBuild");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.PowerPointApi.SlideShowWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			_eventBinding.RaiseCustomEvent("SlideShowNextBuild", ref paramsArray);
		}

		public void SlideShowNextSlide([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SlideShowNextSlide");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.PowerPointApi.SlideShowWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			_eventBinding.RaiseCustomEvent("SlideShowNextSlide", ref paramsArray);
		}

		public void SlideShowEnd([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SlideShowEnd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			_eventBinding.RaiseCustomEvent("SlideShowEnd", ref paramsArray);
		}

		public void PresentationPrint([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PresentationPrint");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			_eventBinding.RaiseCustomEvent("PresentationPrint", ref paramsArray);
		}

		public void SlideSelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SlideSelectionChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sldRange);
				return;
			}

			NetOffice.PowerPointApi.SlideRange newSldRange = Factory.CreateObjectFromComProxy(_eventClass, sldRange) as NetOffice.PowerPointApi.SlideRange;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSldRange;
			_eventBinding.RaiseCustomEvent("SlideSelectionChanged", ref paramsArray);
		}

		public void ColorSchemeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ColorSchemeChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sldRange);
				return;
			}

			NetOffice.PowerPointApi.SlideRange newSldRange = Factory.CreateObjectFromComProxy(_eventClass, sldRange) as NetOffice.PowerPointApi.SlideRange;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSldRange;
			_eventBinding.RaiseCustomEvent("ColorSchemeChanged", ref paramsArray);
		}

		public void PresentationBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PresentationBeforeSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres, cancel);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			object[] paramsArray = new object[2];
			paramsArray[0] = newPres;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("PresentationBeforeSave", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void SlideShowNextClick([In, MarshalAs(UnmanagedType.IDispatch)] object wn, [In, MarshalAs(UnmanagedType.IDispatch)] object nEffect)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SlideShowNextClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn, nEffect);
				return;
			}

			NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.PowerPointApi.SlideShowWindow;
			NetOffice.PowerPointApi.Effect newnEffect = Factory.CreateObjectFromComProxy(_eventClass, nEffect) as NetOffice.PowerPointApi.Effect;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWn;
			paramsArray[1] = newnEffect;
			_eventBinding.RaiseCustomEvent("SlideShowNextClick", ref paramsArray);
		}

		public void AfterNewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterNewPresentation");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			_eventBinding.RaiseCustomEvent("AfterNewPresentation", ref paramsArray);
		}

		public void AfterPresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterPresentationOpen");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			_eventBinding.RaiseCustomEvent("AfterPresentationOpen", ref paramsArray);
		}

		public void PresentationSync([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] object syncEventType)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PresentationSync");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres, syncEventType);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
			object[] paramsArray = new object[2];
			paramsArray[0] = newPres;
			paramsArray[1] = newSyncEventType;
			_eventBinding.RaiseCustomEvent("PresentationSync", ref paramsArray);
		}

		public void SlideShowOnNext([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SlideShowOnNext");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.PowerPointApi.SlideShowWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			_eventBinding.RaiseCustomEvent("SlideShowOnNext", ref paramsArray);
		}

		public void SlideShowOnPrevious([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SlideShowOnPrevious");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.PowerPointApi.SlideShowWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			_eventBinding.RaiseCustomEvent("SlideShowOnPrevious", ref paramsArray);
		}

		public void PresentationBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PresentationBeforeClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres, cancel);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			object[] paramsArray = new object[2];
			paramsArray[0] = newPres;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("PresentationBeforeClose", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowOpen");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(protViewWindow);
				return;
			}

			NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateObjectFromComProxy(_eventClass, protViewWindow) as NetOffice.PowerPointApi.ProtectedViewWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newProtViewWindow;
			_eventBinding.RaiseCustomEvent("ProtectedViewWindowOpen", ref paramsArray);
		}

		public void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowBeforeEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(protViewWindow, cancel);
				return;
			}

			NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateObjectFromComProxy(_eventClass, protViewWindow) as NetOffice.PowerPointApi.ProtectedViewWindow;
			object[] paramsArray = new object[2];
			paramsArray[0] = newProtViewWindow;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeEdit", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] object protectedViewCloseReason, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowBeforeClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(protViewWindow, protectedViewCloseReason, cancel);
				return;
			}

			NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateObjectFromComProxy(_eventClass, protViewWindow) as NetOffice.PowerPointApi.ProtectedViewWindow;
			NetOffice.PowerPointApi.Enums.PpProtectedViewCloseReason newProtectedViewCloseReason = (NetOffice.PowerPointApi.Enums.PpProtectedViewCloseReason)protectedViewCloseReason;
			object[] paramsArray = new object[3];
			paramsArray[0] = newProtViewWindow;
			paramsArray[1] = newProtectedViewCloseReason;
			paramsArray.SetValue(cancel, 2);
			_eventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeClose", ref paramsArray);

			cancel = (bool)paramsArray[2];
		}

		public void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(protViewWindow);
				return;
			}

			NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateObjectFromComProxy(_eventClass, protViewWindow) as NetOffice.PowerPointApi.ProtectedViewWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newProtViewWindow;
			_eventBinding.RaiseCustomEvent("ProtectedViewWindowActivate", ref paramsArray);
		}

		public void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(protViewWindow);
				return;
			}

			NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateObjectFromComProxy(_eventClass, protViewWindow) as NetOffice.PowerPointApi.ProtectedViewWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newProtViewWindow;
			_eventBinding.RaiseCustomEvent("ProtectedViewWindowDeactivate", ref paramsArray);
		}

		public void PresentationCloseFinal([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PresentationCloseFinal");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pres);
				return;
			}

			NetOffice.PowerPointApi.Presentation newPres = Factory.CreateObjectFromComProxy(_eventClass, pres) as NetOffice.PowerPointApi.Presentation;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPres;
			_eventBinding.RaiseCustomEvent("PresentationCloseFinal", ref paramsArray);
		}

		public void AfterDragDropOnSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterDragDropOnSlide");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sld, x, y);
				return;
			}

			NetOffice.PowerPointApi.Slide newSld = Factory.CreateObjectFromComProxy(_eventClass, sld) as NetOffice.PowerPointApi.Slide;
			Single newX = Convert.ToSingle(x);
			Single newY = Convert.ToSingle(y);
			object[] paramsArray = new object[3];
			paramsArray[0] = newSld;
			paramsArray[1] = newX;
			paramsArray[2] = newY;
			_eventBinding.RaiseCustomEvent("AfterDragDropOnSlide", ref paramsArray);
		}

		public void AfterShapeSizeChange([In, MarshalAs(UnmanagedType.IDispatch)] object shp)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterShapeSizeChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shp);
				return;
			}

			NetOffice.PowerPointApi.Shape newshp = Factory.CreateObjectFromComProxy(_eventClass, shp) as NetOffice.PowerPointApi.Shape;
			object[] paramsArray = new object[1];
			paramsArray[0] = newshp;
			_eventBinding.RaiseCustomEvent("AfterShapeSizeChange", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}