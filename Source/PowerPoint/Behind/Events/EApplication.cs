using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.PowerPointApi.EventContracts.EApplication"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class EApplication_SinkHelper : SinkHelper, NetOffice.PowerPointApi.EventContracts.EApplication
    {
        #region Static

        /// <summary>
        /// Interface Id from EApplication
        /// </summary>
        public static readonly string Id = "914934C2-5A91-11CF-8700-00AA0060263B";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public EApplication_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region EApplication

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sel"></param>
        public void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel)
        {
            if (!Validate("WindowSelectionChange"))
            {
                Invoker.ReleaseParamsArray(sel);
                return;
            }

            NetOffice.PowerPointApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Selection>(EventClass, sel, typeof(NetOffice.PowerPointApi.Selection));
            object[] paramsArray = new object[1];
            paramsArray[0] = newSel;
            EventBinding.RaiseCustomEvent("WindowSelectionChange", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="cancel"></param>
        public void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
        {
            if (!Validate("WindowBeforeRightClick"))
            {
                Invoker.ReleaseParamsArray(sel, cancel);
                return;
            }

            NetOffice.PowerPointApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Selection>(EventClass, sel, typeof(NetOffice.PowerPointApi.Selection));
            object[] paramsArray = new object[2];
            paramsArray[0] = newSel;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("WindowBeforeRightClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="cancel"></param>
        public void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
        {
            if (!Validate("WindowBeforeDoubleClick"))
            {
                Invoker.ReleaseParamsArray(sel, cancel);
                return;
            }

            NetOffice.PowerPointApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Selection>(EventClass, sel, typeof(NetOffice.PowerPointApi.Selection));
            object[] paramsArray = new object[2];
            paramsArray[0] = newSel;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("WindowBeforeDoubleClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        public void PresentationClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("PresentationClose"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPres;
            EventBinding.RaiseCustomEvent("PresentationClose", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        public void PresentationSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("PresentationSave"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPres;
            EventBinding.RaiseCustomEvent("PresentationSave", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        public void PresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("PresentationOpen"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPres;
            EventBinding.RaiseCustomEvent("PresentationOpen", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        public void NewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("NewPresentation"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPres;
            EventBinding.RaiseCustomEvent("NewPresentation", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sld"></param>
        public void PresentationNewSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld)
        {
            if (!Validate("PresentationNewSlide"))
            {
                Invoker.ReleaseParamsArray(sld);
                return;
            }

            NetOffice.PowerPointApi.Slide newSld = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Slide>(EventClass, sld, typeof(NetOffice.PowerPointApi.Slide));
            object[] paramsArray = new object[1];
            paramsArray[0] = newSld;
            EventBinding.RaiseCustomEvent("PresentationNewSlide", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        /// <param name="wn"></param>
        public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowActivate"))
            {
                Invoker.ReleaseParamsArray(pres, wn);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            NetOffice.PowerPointApi.DocumentWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.DocumentWindow>(EventClass, wn, typeof(NetOffice.PowerPointApi.DocumentWindow));
            object[] paramsArray = new object[2];
            paramsArray[0] = newPres;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        /// <param name="wn"></param>
        public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(pres, wn);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            NetOffice.PowerPointApi.DocumentWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.DocumentWindow>(EventClass, wn, typeof(NetOffice.PowerPointApi.DocumentWindow));
            object[] paramsArray = new object[2];
            paramsArray[0] = newPres;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wn"></param>
        public void SlideShowBegin([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("SlideShowBegin"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.PowerPointApi.DocumentWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.DocumentWindow>(EventClass, wn, typeof(NetOffice.PowerPointApi.DocumentWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWn;
            EventBinding.RaiseCustomEvent("SlideShowBegin", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wn"></param>
        public void SlideShowNextBuild([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("SlideShowNextBuild"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.PowerPointApi.DocumentWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.DocumentWindow>(EventClass, wn, typeof(NetOffice.PowerPointApi.DocumentWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWn;
            EventBinding.RaiseCustomEvent("SlideShowNextBuild", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wn"></param>
        public void SlideShowNextSlide([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("SlideShowNextSlide"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.PowerPointApi.DocumentWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.DocumentWindow>(EventClass, wn, typeof(NetOffice.PowerPointApi.DocumentWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWn;
            EventBinding.RaiseCustomEvent("SlideShowNextSlide", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        public void SlideShowEnd([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("SlideShowEnd"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPres;
            EventBinding.RaiseCustomEvent("SlideShowEnd", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        public void PresentationPrint([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("PresentationPrint"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPres;
            EventBinding.RaiseCustomEvent("PresentationPrint", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sldRange"></param>
        public void SlideSelectionChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange)
        {
            if (!Validate("SlideSelectionChanged"))
            {
                Invoker.ReleaseParamsArray(sldRange);
                return;
            }

            NetOffice.PowerPointApi.SlideRange newSldRange = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideRange>(EventClass, sldRange, typeof(NetOffice.PowerPointApi.SlideRange));
            object[] paramsArray = new object[1];
            paramsArray[0] = newSldRange;
            EventBinding.RaiseCustomEvent("SlideSelectionChanged", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sldRange"></param>
        public void ColorSchemeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object sldRange)
        {
            if (!Validate("ColorSchemeChanged"))
            {
                Invoker.ReleaseParamsArray(sldRange);
                return;
            }

            NetOffice.PowerPointApi.SlideRange newSldRange = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideRange>(EventClass, sldRange, typeof(NetOffice.PowerPointApi.SlideRange));
            object[] paramsArray = new object[1];
            paramsArray[0] = newSldRange;
            EventBinding.RaiseCustomEvent("ColorSchemeChanged", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        /// <param name="cancel"></param>
        public void PresentationBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel)
        {
            if (!Validate("PresentationBeforeSave"))
            {
                Invoker.ReleaseParamsArray(pres, cancel);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            object[] paramsArray = new object[2];
            paramsArray[0] = newPres;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("PresentationBeforeSave", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wn"></param>
        /// <param name="nEffect"></param>
        public void SlideShowNextClick([In, MarshalAs(UnmanagedType.IDispatch)] object wn, [In, MarshalAs(UnmanagedType.IDispatch)] object nEffect)
        {
            if (!Validate("SlideShowNextClick"))
            {
                Invoker.ReleaseParamsArray(wn, nEffect);
                return;
            }

            NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideShowWindow>(EventClass, wn, typeof(NetOffice.PowerPointApi.SlideShowWindow));
            NetOffice.PowerPointApi.Effect newnEffect = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Effect>(EventClass, nEffect, typeof(NetOffice.PowerPointApi.Effect));
            object[] paramsArray = new object[2];
            paramsArray[0] = newWn;
            paramsArray[1] = newnEffect;
            EventBinding.RaiseCustomEvent("SlideShowNextClick", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        public void AfterNewPresentation([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("AfterNewPresentation"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPres;
            EventBinding.RaiseCustomEvent("AfterNewPresentation", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        public void AfterPresentationOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("AfterPresentationOpen"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPres;
            EventBinding.RaiseCustomEvent("AfterPresentationOpen", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        /// <param name="syncEventType"></param>
        public void PresentationSync([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] object syncEventType)
        {
            if (!Validate("PresentationSync"))
            {
                Invoker.ReleaseParamsArray(pres, syncEventType);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
            object[] paramsArray = new object[2];
            paramsArray[0] = newPres;
            paramsArray[1] = newSyncEventType;
            EventBinding.RaiseCustomEvent("PresentationSync", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wn"></param>
        public void SlideShowOnNext([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("SlideShowOnNext"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideShowWindow>(EventClass, wn, typeof(NetOffice.PowerPointApi.SlideShowWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWn;
            EventBinding.RaiseCustomEvent("SlideShowOnNext", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wn"></param>
        public void SlideShowOnPrevious([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("SlideShowOnPrevious"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.PowerPointApi.SlideShowWindow newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.SlideShowWindow>(EventClass, wn, typeof(NetOffice.PowerPointApi.SlideShowWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newWn;
            EventBinding.RaiseCustomEvent("SlideShowOnPrevious", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        /// <param name="cancel"></param>
        public void PresentationBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pres, [In] [Out] ref object cancel)
        {
            if (!Validate("SlideShowOnPrevious"))
            {
                Invoker.ReleaseParamsArray(pres, cancel);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            object[] paramsArray = new object[2];
            paramsArray[0] = newPres;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("PresentationBeforeClose", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="protViewWindow"></param>
        public void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow)
        {
            if (!Validate("ProtectedViewWindowOpen"))
            {
                Invoker.ReleaseParamsArray(protViewWindow);
                return;
            }

            NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.ProtectedViewWindow>(EventClass, protViewWindow, typeof(NetOffice.PowerPointApi.ProtectedViewWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newProtViewWindow;
            EventBinding.RaiseCustomEvent("ProtectedViewWindowOpen", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="protViewWindow"></param>
        /// <param name="cancel"></param>
        public void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] [Out] ref object cancel)
        {
            if (!Validate("ProtectedViewWindowBeforeEdit"))
            {
                Invoker.ReleaseParamsArray(protViewWindow, cancel);
                return;
            }

            NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.ProtectedViewWindow>(EventClass, protViewWindow, typeof(NetOffice.PowerPointApi.ProtectedViewWindow));
            object[] paramsArray = new object[2];
            paramsArray[0] = newProtViewWindow;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeEdit", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="protViewWindow"></param>
        /// <param name="protectedViewCloseReason"></param>
        /// <param name="cancel"></param>
        public void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow, [In] object protectedViewCloseReason, [In] [Out] ref object cancel)
        {
            if (!Validate("ProtectedViewWindowBeforeClose"))
            {
                Invoker.ReleaseParamsArray(protViewWindow, protectedViewCloseReason, cancel);
                return;
            }

            NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.ProtectedViewWindow>(EventClass, protViewWindow, typeof(NetOffice.PowerPointApi.ProtectedViewWindow));
            NetOffice.PowerPointApi.Enums.PpProtectedViewCloseReason newProtectedViewCloseReason = (NetOffice.PowerPointApi.Enums.PpProtectedViewCloseReason)protectedViewCloseReason;
            object[] paramsArray = new object[3];
            paramsArray[0] = newProtViewWindow;
            paramsArray[1] = newProtectedViewCloseReason;
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("ProtectedViewWindowBeforeClose", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="protViewWindow"></param>
        public void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow)
        {
            if (!Validate("ProtectedViewWindowActivate"))
            {
                Invoker.ReleaseParamsArray(protViewWindow);
                return;
            }

            NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.ProtectedViewWindow>(EventClass, protViewWindow, typeof(NetOffice.PowerPointApi.ProtectedViewWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newProtViewWindow;
            EventBinding.RaiseCustomEvent("ProtectedViewWindowActivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="protViewWindow"></param>
        public void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object protViewWindow)
        {
            if (!Validate("ProtectedViewWindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(protViewWindow);
                return;
            }

            NetOffice.PowerPointApi.ProtectedViewWindow newProtViewWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.ProtectedViewWindow>(EventClass, protViewWindow, typeof(NetOffice.PowerPointApi.ProtectedViewWindow));
            object[] paramsArray = new object[1];
            paramsArray[0] = newProtViewWindow;
            EventBinding.RaiseCustomEvent("ProtectedViewWindowDeactivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pres"></param>
        public void PresentationCloseFinal([In, MarshalAs(UnmanagedType.IDispatch)] object pres)
        {
            if (!Validate("PresentationCloseFinal"))
            {
                Invoker.ReleaseParamsArray(pres);
                return;
            }

            NetOffice.PowerPointApi.Presentation newPres = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Presentation>(EventClass, pres, typeof(NetOffice.PowerPointApi.Presentation));
            object[] paramsArray = new object[1];
            paramsArray[0] = newPres;
            EventBinding.RaiseCustomEvent("PresentationCloseFinal", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sld"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public void AfterDragDropOnSlide([In, MarshalAs(UnmanagedType.IDispatch)] object sld, [In] object x, [In] object y)
        {
            if (!Validate("AfterDragDropOnSlide"))
            {
                Invoker.ReleaseParamsArray(sld, x, y);
                return;
            }

            NetOffice.PowerPointApi.Slide newSld = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Slide>(EventClass, sld, typeof(NetOffice.PowerPointApi.Slide));
            Single newX = ToSingle(x);
            Single newY = ToSingle(y);
            object[] paramsArray = new object[3];
            paramsArray[0] = newSld;
            paramsArray[1] = newX;
            paramsArray[2] = newY;
            EventBinding.RaiseCustomEvent("AfterDragDropOnSlide", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="shp"></param>
        public void AfterShapeSizeChange([In, MarshalAs(UnmanagedType.IDispatch)] object shp)
        {
            if (!Validate("AfterShapeSizeChange"))
            {
                Invoker.ReleaseParamsArray(shp);
                return;
            }

            NetOffice.PowerPointApi.Shape newshp = Factory.CreateKnownObjectFromComProxy<NetOffice.PowerPointApi.Shape>(EventClass, shp, typeof(NetOffice.PowerPointApi.Shape));
            object[] paramsArray = new object[1];
            paramsArray[0] = newshp;
            EventBinding.RaiseCustomEvent("AfterShapeSizeChange", ref paramsArray);
        }

        #endregion
    }
}

