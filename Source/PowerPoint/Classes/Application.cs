using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PowerPointApi
{
    #region Delegates

    #pragma warning disable
    public delegate void Application_WindowSelectionChangeEventHandler(NetOffice.PowerPointApi.Selection sel);
    public delegate void Application_WindowBeforeRightClickEventHandler(NetOffice.PowerPointApi.Selection sel, ref bool cancel);
    public delegate void Application_WindowBeforeDoubleClickEventHandler(NetOffice.PowerPointApi.Selection sel, ref bool cancel);
    public delegate void Application_PresentationCloseEventHandler(NetOffice.PowerPointApi.Presentation pres);
    public delegate void Application_PresentationSaveEventHandler(NetOffice.PowerPointApi.Presentation pres);
    public delegate void Application_PresentationOpenEventHandler(NetOffice.PowerPointApi.Presentation pres);
    public delegate void Application_NewPresentationEventHandler(NetOffice.PowerPointApi.Presentation pres);
    public delegate void Application_PresentationNewSlideEventHandler(NetOffice.PowerPointApi.Slide sld);
    public delegate void Application_WindowActivateEventHandler(NetOffice.PowerPointApi.Presentation pres, NetOffice.PowerPointApi.DocumentWindow wn);
    public delegate void Application_WindowDeactivateEventHandler(NetOffice.PowerPointApi.Presentation pres, NetOffice.PowerPointApi.DocumentWindow wn);
    public delegate void Application_SlideShowBeginEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn);
    public delegate void Application_SlideShowNextBuildEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn);
    public delegate void Application_SlideShowNextSlideEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn);
    public delegate void Application_SlideShowEndEventHandler(NetOffice.PowerPointApi.Presentation pres);
    public delegate void Application_PresentationPrintEventHandler(NetOffice.PowerPointApi.Presentation pres);
    public delegate void Application_SlideSelectionChangedEventHandler(NetOffice.PowerPointApi.SlideRange sldRange);
    public delegate void Application_ColorSchemeChangedEventHandler(NetOffice.PowerPointApi.SlideRange sldRange);
    public delegate void Application_PresentationBeforeSaveEventHandler(NetOffice.PowerPointApi.Presentation pres, ref bool Cancel);
    public delegate void Application_SlideShowNextClickEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn, NetOffice.PowerPointApi.Effect nEffect);
    public delegate void Application_AfterNewPresentationEventHandler(NetOffice.PowerPointApi.Presentation pres);
    public delegate void Application_AfterPresentationOpenEventHandler(NetOffice.PowerPointApi.Presentation pres);
    public delegate void Application_PresentationSyncEventHandler(NetOffice.PowerPointApi.Presentation pres, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);
    public delegate void Application_SlideShowOnNextEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn);
    public delegate void Application_SlideShowOnPreviousEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn);
    public delegate void Application_PresentationBeforeCloseEventHandler(NetOffice.PowerPointApi.Presentation pres, ref bool cancel);
    public delegate void Application_ProtectedViewWindowOpenEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow protViewWindow);
    public delegate void Application_ProtectedViewWindowBeforeEditEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow protViewWindow, ref bool Cancel);
    public delegate void Application_ProtectedViewWindowBeforeCloseEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow protViewWindow, NetOffice.PowerPointApi.Enums.PpProtectedViewCloseReason protectedViewCloseReason, ref bool cancel);
    public delegate void Application_ProtectedViewWindowActivateEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow protViewWindow);
    public delegate void Application_ProtectedViewWindowDeactivateEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow protViewWindow);
    public delegate void Application_PresentationCloseFinalEventHandler(NetOffice.PowerPointApi.Presentation pres);
    public delegate void Application_AfterDragDropOnSlideEventHandler(NetOffice.PowerPointApi.Slide sld, Single x, Single yY);
    public delegate void Application_AfterShapeSizeChangeEventHandler(NetOffice.PowerPointApi.Shape shp);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Application
    /// This class is an alias/typedef for NetOffice.PowerPointApi.Behind.Application
    /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194565.aspx </remarks>
    [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [InteropCompatibilityClass]
    public class ApplicationClass : NetOffice.PowerPointApi.Behind.Application
    {
        private string _defaultProgId = "PowerPoint.Application";

        /// <summary>
        /// Creates a new instance of Microsoft PowerPoint
        /// </summary>
        public ApplicationClass()
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(_defaultProgId);
        }

        /// <summary>
        /// Creates a new instance of Microsoft PowerPoint based on given id.
        /// This can be used to target a specific version of Microsoft PowerPoint.
        /// Example usage:
        /// "Microsoft.PowerPoint.12" to target PowerPoint 2007
        /// "Microsoft.PowerPoint.14" to target PowerPoint 2010
        /// </summary>
        /// <param name="progId">given progid for specific version</param>
        public ApplicationClass(string progId)
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(progId);
        }

        /// <summary>
        /// Try get accessing a running application or create a new instance of Microsoft PowerPoint
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public ApplicationClass(Core factory = null, bool tryProxyServiceFirst = false) : base(factory, tryProxyServiceFirst)
        {

        }

        /// <summary>
        /// Creates a new instance of Microsoft PowerPoint
        /// </summary>
        /// <param name="mode">indicates where is the call coming from</param>
        public ApplicationClass(NetOffice.Callers.InteropCompatibilityClassCreateMode mode)
        {
            if (mode == NetOffice.Callers.InteropCompatibilityClassCreateMode.Direct)
            {
                ICOMObjectInitialize init = (ICOMObjectInitialize)this;
                init.InitializeCOMObject(_defaultProgId);
            }
        }
    }

    /// <summary>
    /// CoClass Application
    /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745704.aspx </remarks>
    [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass), ComProgId("PowerPoint.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(EventContracts.EApplication))]
	[TypeId("91493441-5A91-11CF-8700-00AA0060263B")]
    public interface Application : _Application, ICloneable<Application>, IEventBinding, ICOMObjectProxyService
    {
        #region Events

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743918.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowSelectionChangeEventHandler WindowSelectionChangeEvent;
        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746559.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowBeforeRightClickEventHandler WindowBeforeRightClickEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745746.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowBeforeDoubleClickEventHandler WindowBeforeDoubleClickEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744678.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_PresentationCloseEventHandler PresentationCloseEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744230.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_PresentationSaveEventHandler PresentationSaveEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744100.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_PresentationOpenEventHandler PresentationOpenEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745073.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_NewPresentationEventHandler NewPresentationEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746597.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_PresentationNewSlideEventHandler PresentationNewSlideEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743995.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowActivateEventHandler WindowActivateEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745519.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_WindowDeactivateEventHandler WindowDeactivateEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746741.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SlideShowBeginEventHandler SlideShowBeginEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745070.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SlideShowNextBuildEventHandler SlideShowNextBuildEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745863.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SlideShowNextSlideEventHandler SlideShowNextSlideEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746536.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_SlideShowEndEventHandler SlideShowEndEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744696.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event Application_PresentationPrintEventHandler PresentationPrintEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745869.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        event Application_SlideSelectionChangedEventHandler SlideSelectionChangedEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745549.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        event Application_ColorSchemeChangedEventHandler ColorSchemeChangedEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744682.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        event Application_PresentationBeforeSaveEventHandler PresentationBeforeSaveEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745682.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        event Application_SlideShowNextClickEventHandler SlideShowNextClickEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746421.aspx </remarks>
        [SupportByVersion("PowerPoint", 11, 12, 14, 15, 16)]
        event Application_AfterNewPresentationEventHandler AfterNewPresentationEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744659.aspx </remarks>
        [SupportByVersion("PowerPoint", 11, 12, 14, 15, 16)]
        event Application_AfterPresentationOpenEventHandler AfterPresentationOpenEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744576.aspx </remarks>
        [SupportByVersion("PowerPoint", 11, 12, 14, 15, 16)]
        event Application_PresentationSyncEventHandler PresentationSyncEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746469.aspx </remarks>
        [SupportByVersion("PowerPoint", 12, 14, 15, 16)]
        event Application_SlideShowOnNextEventHandler SlideShowOnNextEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744749.aspx </remarks>
        [SupportByVersion("PowerPoint", 12, 14, 15, 16)]
        event Application_SlideShowOnPreviousEventHandler SlideShowOnPreviousEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745567.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        event Application_PresentationBeforeCloseEventHandler PresentationBeforeCloseEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745081.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        event Application_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpenEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745575.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        event Application_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEditEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746497.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        event Application_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeCloseEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744591.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        event Application_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivateEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746253.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        event Application_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivateEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744781.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        event Application_PresentationCloseFinalEventHandler PresentationCloseFinalEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227644.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        event Application_AfterDragDropOnSlideEventHandler AfterDragDropOnSlideEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227375.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        event Application_AfterShapeSizeChangeEventHandler AfterShapeSizeChangeEvent;

        #endregion
    }
}
